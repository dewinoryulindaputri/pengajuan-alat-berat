from flask import Flask, render_template, request, redirect, url_for, session
import os
import io
import pandas as pd
import openpyxl
from flask import send_file
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2 import service_account
from pypdf import PdfReader 
from xhtml2pdf import pisa   
from werkzeug.utils import secure_filename 
import sqlite3
import json

app = Flask(__name__)
app.secret_key = 'kunci_rahasia_spip'

# --- KONFIGURASI GOOGLE DRIVE ---
SERVICE_ACCOUNT_FILE = 'client_secret.json'
SCOPES = ['https://www.googleapis.com/auth/drive']
GOOGLE_DRIVE_FOLDER_ID = '1H8crWdaeqPcUHrlpqaqrs4F2WoY89nzg'

def get_drive_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def get_or_create_subfolder(service, parent_folder_id, folder_name):
    try:
        proper_name = folder_name.strip().capitalize()
        query = f"name = '{proper_name}' and '{parent_folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        folders = results.get('files', [])
        if folders:
            return folders[0]['id']  
        else:
            folder_metadata = {
                'name': proper_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_folder_id]
            }
            folder = service.files().create(body=folder_metadata, fields='id').execute()
            return folder.get('id')
    except Exception as e:
        print(f"Gagal membuat/mencari subfolder di Drive: {e}")
        return parent_folder_id 

# --- FUNGSI: Mendeteksi kategori/nama folder berdasarkan isi PDF ---
def detect_category_from_pdf(file_storage):
    try:
        file_storage.seek(0)
        reader = PdfReader(file_storage)
        text = ""
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted.lower()
        
        file_storage.seek(0) 
        
        if "sarana" in text:
            return "Sarana"
        elif "prasarana" in text:
            return "Prasarana"
        elif "instalasi" in text:
            return "Instalasi"
        elif "peralatan" in text:
            return "Peralatan"
        else:
            return "Laporan_Umum"
    except Exception as e:
        print(f"Gagal membaca teks PDF untuk deteksi folder: {e}")
        file_storage.seek(0)
        return "Laporan_Lainnya"

def upload_to_drive(file_storage, kategori):
    try:
        service = get_drive_service()
        
        # Selalu pakai kategori yang dipilih user di form (sarana/prasarana/instalasi/peralatan)
        # sebagai nama folder tujuan, baik file berupa PDF, foto, maupun tipe lainnya.
        # Ini lebih akurat dibanding menebak dari isi teks PDF, apalagi PDF sering
        # berisi foto/scan yang tidak punya teks untuk dibaca.
        nama_subfolder = kategori.capitalize()
            
        target_folder_id = get_or_create_subfolder(service, GOOGLE_DRIVE_FOLDER_ID, nama_subfolder)
        
        file_storage.seek(0)
        media_body = MediaIoBaseUpload(file_storage, mimetype=file_storage.content_type, resumable=True)
        
        file_metadata = {
            'name': file_storage.filename, 
            'parents': [target_folder_id]
        }
        
        file = service.files().create(body=file_metadata, media_body=media_body, fields='id').execute()
        return file.get('id')
    except Exception as e:
        print(f"DEBUG ERROR UPLOAD: {e}")
        return None

# --- Fungsi untuk generate dan upload laporan PDF ke Drive ---
def generate_and_upload_laporan_pdf(list_data):
    try:
        html_content = "<html><body><h2>Laporan Rekapitulasi Perizinan Alat Berat</h2><table border='1' cellspacing='0' cellpadding='5'><tr><th>ID</th><th>Kode</th><th>Pemohon</th><th>Kategori</th><th>Status</th></tr>"
        for item in list_data:
            html_content += f"<tr><td>{item['id']}</td><td>{item['kode_unik']}</td><td>{item['pemohon']}</td><td>{item['kategori']}</td><td>{item['status']}</td></tr>"
        html_content += "</table></body></html>"

        pdf_file = io.BytesIO()
        pisa_status = pisa.CreatePDF(io.BytesIO(html_content.encode('utf-8')), dest=pdf_file)
        if pisa_status.err:
            return None
        pdf_file.seek(0)

        service = get_drive_service()
        target_folder_id = get_or_create_subfolder(service, GOOGLE_DRIVE_FOLDER_ID, "Rekap_Laporan")
        media_body = MediaIoBaseUpload(pdf_file, mimetype='application/pdf', resumable=True)
        file_metadata = {
            'name': f'Rekap_Laporan_{pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")}.pdf', 
            'parents': [target_folder_id]
        }
        
        file = service.files().create(body=file_metadata, media_body=media_body, fields='id').execute()
        return file.get('id')
    except Exception as e:
        print(f"Error saat generate/upload PDF Laporan: {e}")
        return None

UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- KONFIGURASI DATABASE SQLITE ---
# Kalau folder /app/data ada (artinya Volume Railway sudah dipasang & di-mount di situ),
# simpan database.db di sana supaya persisten antar redeploy.
# Kalau tidak ada (misal saat dijalankan di laptop/lokal), tetap pakai folder project seperti biasa.
if os.path.isdir('/app/data'):
    DB_FILE = '/app/data/database.db'
else:
    DB_FILE = 'database.db'

def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS pengguna (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT NOT NULL,
            email TEXT,
            role TEXT
        )
    ''')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS permohonan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kode_unik TEXT, pemohon TEXT, kategori TEXT,
            jenis_sarana TEXT, jenis_prasarana TEXT, jenis_instalasi TEXT, jenis_peralatan TEXT,
            tahun_pembuatan TEXT, kapasitas_orang TEXT, no_lambung TEXT, no_polisi TEXT,
            merk TEXT, tipe TEXT, nomer_mesin TEXT, nomer_rangka TEXT, nomer_stnk TEXT,
            instansi TEXT, nama_prasarana TEXT, tahun_konstruksi TEXT, lokasi TEXT, koordinat TEXT,
            tahun TEXT, kapasitas TEXT, no_sertifikat TEXT, tgl_berlaku TEXT, tgl_berakhir TEXT,
            catatan TEXT, status TEXT, foto TEXT, foto_drive_ids TEXT
        )
    ''')
    if conn.execute('SELECT COUNT(*) FROM pengguna').fetchone()[0] == 0:
        conn.execute("INSERT INTO pengguna (nama, email, role) VALUES (?, ?, ?)",
                     ('Super Administrator', 'admin@perizinan.com', 'Administrator'))
        conn.execute("INSERT INTO pengguna (nama, email, role) VALUES (?, ?, ?)",
                     ('Petugas Verifikasi', 'petugas@perizinan.com', 'Petugas'))
    conn.commit()
    conn.close()

def get_all_pengguna():
    conn = get_db_connection()
    rows = conn.execute('SELECT * FROM pengguna').fetchall()
    conn.close()
    return [dict(r) for r in rows]

def cek_pengguna_ada(nama):
    conn = get_db_connection()
    row = conn.execute('SELECT id FROM pengguna WHERE LOWER(nama) = LOWER(?)', (nama,)).fetchone()
    conn.close()
    return row is not None

def add_pengguna(nama, email, role):
    conn = get_db_connection()
    conn.execute('INSERT INTO pengguna (nama, email, role) VALUES (?, ?, ?)', (nama, email, role))
    conn.commit()
    conn.close()

def _row_to_permohonan(row):
    item = dict(row)
    item['foto'] = json.loads(item['foto']) if item['foto'] else []
    item['foto_drive_ids'] = json.loads(item['foto_drive_ids']) if item['foto_drive_ids'] else []
    return item

def get_all_permohonan():
    conn = get_db_connection()
    rows = conn.execute('SELECT * FROM permohonan ORDER BY id').fetchall()
    conn.close()
    return [_row_to_permohonan(r) for r in rows]

def get_sarana_disetujui():
    conn = get_db_connection()
    rows = conn.execute(
        "SELECT * FROM permohonan WHERE kategori = 'sarana' AND status = 'Disetujui' ORDER BY id"
    ).fetchall()
    conn.close()
    hasil = []
    for r in rows:
        d = dict(r)
        hasil.append({
            'id': d['id'],
            'jenis_sarana': d['jenis_sarana'],
            'no_lambung': d['no_lambung'],
            'merk': d['merk'],
            'tipe': d['tipe'],
            'instansi': d['instansi'],
            'status': 'Aktif'
        })
    return hasil

def get_prasarana_disetujui():
    conn = get_db_connection()
    rows = conn.execute(
        "SELECT * FROM permohonan WHERE kategori = 'prasarana' AND status = 'Disetujui' ORDER BY id"
    ).fetchall()
    conn.close()
    hasil = []
    for r in rows:
        d = dict(r)
        hasil.append({
            'id': d['id'],
            'nama_prasarana': d['nama_prasarana'],
            'lokasi': d['lokasi'],
            'status': 'Layak'
        })
    return hasil

def get_instalasi_disetujui():
    conn = get_db_connection()
    rows = conn.execute(
        "SELECT * FROM permohonan WHERE kategori = 'instalasi' AND status = 'Disetujui' ORDER BY id"
    ).fetchall()
    conn.close()
    hasil = []
    for r in rows:
        d = dict(r)
        hasil.append({
            'id': d['id'],
            'jenis_instalasi': d['jenis_instalasi'],
            'kapasitas': d['kapasitas'],
            'status': 'Beroperasi'
        })
    return hasil

def get_peralatan_disetujui():
    conn = get_db_connection()
    rows = conn.execute(
        "SELECT * FROM permohonan WHERE kategori = 'peralatan' AND status = 'Disetujui' ORDER BY id"
    ).fetchall()
    conn.close()
    hasil = []
    for r in rows:
        d = dict(r)
        hasil.append({
            'id': d['id'],
            'jenis_peralatan': d['jenis_peralatan'],
            'catatan': d['catatan'],
            'status': 'Normal'
        })
    return hasil

def get_permohonan_by_id(id_permohonan):
    conn = get_db_connection()
    row = conn.execute('SELECT * FROM permohonan WHERE id = ?', (id_permohonan,)).fetchone()
    conn.close()
    return _row_to_permohonan(row) if row else None

def get_permohonan_by_pemohon(nama_pemohon):
    conn = get_db_connection()
    rows = conn.execute('SELECT * FROM permohonan WHERE pemohon = ? ORDER BY id', (nama_pemohon,)).fetchall()
    conn.close()
    return [_row_to_permohonan(r) for r in rows]

def add_permohonan(data):
    conn = get_db_connection()
    cursor = conn.execute('''
        INSERT INTO permohonan (
            kode_unik, pemohon, kategori, jenis_sarana, jenis_prasarana, jenis_instalasi, jenis_peralatan,
            tahun_pembuatan, kapasitas_orang, no_lambung, no_polisi, merk, tipe, nomer_mesin, nomer_rangka,
            nomer_stnk, instansi, nama_prasarana, tahun_konstruksi, lokasi, koordinat, tahun, kapasitas,
            no_sertifikat, tgl_berlaku, tgl_berakhir, catatan, status, foto, foto_drive_ids
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        'TEMP', data['pemohon'], data['kategori'], data['jenis_sarana'], data['jenis_prasarana'],
        data['jenis_instalasi'], data['jenis_peralatan'], data['tahun_pembuatan'], data['kapasitas_orang'],
        data['no_lambung'], data['no_polisi'], data['merk'], data['tipe'], data['nomer_mesin'],
        data['nomer_rangka'], data['nomer_stnk'], data['instansi'], data['nama_prasarana'],
        data['tahun_konstruksi'], data['lokasi'], data['koordinat'], data['tahun'], data['kapasitas'],
        data['no_sertifikat'], data['tgl_berlaku'], data['tgl_berakhir'], data['catatan'], data['status'],
        json.dumps(data['foto']), json.dumps(data['foto_drive_ids'])
    ))
    new_id = cursor.lastrowid
    conn.execute('UPDATE permohonan SET kode_unik = ? WHERE id = ?', (f"REQ-{new_id:03d}", new_id))
    conn.commit()
    conn.close()

def update_status_permohonan(id_permohonan, status):
    conn = get_db_connection()
    conn.execute('UPDATE permohonan SET status = ? WHERE id = ?', (status, id_permohonan))
    conn.commit()
    conn.close()

init_db()

data_sarana = [{'id': 1, 'jenis_sarana': 'Light Vehicle', 'no_lambung': 'DT-001', 'merk': 'Scania', 'tipe': 'P460', 'instansi': 'Departemen Tambang', 'status': 'Aktif'}]

data_prasarana = [
    {'id': 1, 'nama_prasarana': 'Gudang Logistik Utama', 'lokasi': 'Kawasan Industri Blok A', 'status': 'Layak'},
    {'id': 2, 'nama_prasarana': 'Workshop Perbaikan Alat Berat', 'lokasi': 'Zona Tambang B', 'status': 'Layak'}
]

data_instalasi = [
    {'id': 1, 'jenis_instalasi': 'Instalasi Pengolahan Air (IPA)', 'kapasitas': '50 Liter/detik', 'status': 'Beroperasi'}
]

data_peralatan = [
    {'id': 1, 'jenis_peralatan': 'Genset 500 KVA', 'catatan': 'Pembangkit Listrik Cadangan', 'status': 'Normal'}
]

@app.route('/')
def login_page():
    return render_template('login.html')

@app.route('/login', methods=['POST'])
def do_login():
    username = request.form.get('username', '').strip() or request.form.get('nama', '').strip()
    password = request.form.get('password', '').strip()
    session.clear() 
    
    if username.lower() == 'admin' and password == '12345':
        session['user'] = 'admin'
        session['nama'] = 'Administrator'
    else:
        nama_aktif = username if username else 'User'
        session['user'] = 'user'
        session['nama'] = nama_aktif
        
        if not cek_pengguna_ada(nama_aktif):
            add_pengguna(
                nama_aktif,
                f"{nama_aktif.lower().replace(' ', '')}@perizinan.com",
                'User Pemohon'
            )
            
    return redirect(url_for('dashboard_page'))

@app.route('/dashboard')
def dashboard_page():
    role = session.get('user')
    if role == 'admin':
        return redirect(url_for('admin_page'))
    elif role == 'user':
        current_user = session.get('nama', 'User')
        user_permohonan = get_permohonan_by_pemohon(current_user)
        total_diajukan = len(user_permohonan)
        total_menunggu = sum(1 for item in user_permohonan if item['status'] == 'Pending')
        total_disetujui = sum(1 for item in user_permohonan if item['status'] == 'Disetujui')
        total_ditolak = sum(1 for item in user_permohonan if item['status'] == 'Ditolak')
        
        permohonan_terbaru = user_permohonan[::-1][:5]
        formatted_terbaru = []
        for idx, item in enumerate(permohonan_terbaru):
            nama_alat = (
                item.get('jenis_sarana') or 
                item.get('jenis_prasarana') or 
                item.get('jenis_instalasi') or 
                item.get('jenis_peralatan') or 
                'Unit / Alat'
            )
            formatted_terbaru.append({
                'no': idx + 1,
                'nama_alat': f"{item.get('kode_unik', 'REQ')} - {nama_alat}",
                'tanggal': 'Baru saja',
                'status': 'Disetujui' if item['status'] == 'Disetujui' else ('Ditolak' if item['status'] == 'Ditolak' else 'Menunggu')
            })

        return render_template(
            'index.html',
            total_diajukan=total_diajukan,
            total_menunggu=total_menunggu,
            total_disetujui=total_disetujui,
            total_ditolak=total_ditolak,
            permohonan_terbaru=formatted_terbaru
        )
    return redirect(url_for('login_page'))

@app.route('/admin')
def admin_page():
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    return render_template('admin.html', list_permohonan=get_all_permohonan())

@app.route('/permohonan-masuk')
def permohonan_masuk_page():
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    return render_template('permohonan_masuk.html', list_permohonan=get_all_permohonan())

@app.route('/detail-permohonan/<int:id_permohonan>')
def detail_permohonan_page(id_permohonan):
    if not session.get('user'):
        return redirect(url_for('login_page'))
    permohonan = get_permohonan_by_id(id_permohonan)
    if not permohonan:
        return "Data permohonan tidak ditemukan", 404
    return render_template('detail_permohonan.html', permohonan=permohonan)

@app.route('/setujui/<int:id_permohonan>')
def setujui_permohonan(id_permohonan):
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    update_status_permohonan(id_permohonan, 'Disetujui')
    return redirect(url_for('permohonan_masuk_page'))

@app.route('/tolak/<int:id_permohonan>')
def tolak_permohonan(id_permohonan):
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    update_status_permohonan(id_permohonan, 'Ditolak')
    return redirect(url_for('permohonan_masuk_page'))

@app.route('/master-sarana')
def master_sarana_page():
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    gabungan_sarana = data_sarana + get_sarana_disetujui()
    return render_template('master_sarana.html', list_sarana=gabungan_sarana)

@app.route('/master-prasarana')
def master_prasarana_page():
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    gabungan_prasarana = data_prasarana + get_prasarana_disetujui()
    return render_template('master_prasarana.html', list_prasarana=gabungan_prasarana)

@app.route('/master-instalasi')
def master_instalasi_page():
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    gabungan_instalasi = data_instalasi + get_instalasi_disetujui()
    return render_template('master_instalasi.html', list_instalasi=gabungan_instalasi)

@app.route('/master-peralatan')
def master_peralatan_page():
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    gabungan_peralatan = data_peralatan + get_peralatan_disetujui()
    return render_template('master_peralatan.html', list_peralatan=gabungan_peralatan)

@app.route('/laporan')
def laporan_page():
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    jenis_filter = request.args.get('jenis', 'Semua Jenis')
    status_filter = request.args.get('status', 'Semua Status')
    filtered_data = get_all_permohonan()
    if jenis_filter != 'Semua Jenis':
        kategori_kunci = jenis_filter.split(' ')[0].lower() 
        filtered_data = [item for item in filtered_data if item.get('kategori', '').lower() == kategori_kunci]
    if status_filter != 'Semua Status':
        mapping_status = {
            'Layak Digunakan': 'Disetujui',
            'Tidak Layak Digunakan': 'Ditolak',
            'Menunggu Verifikasi': 'Pending'
        }
        status_target = mapping_status.get(status_filter)
        filtered_data = [item for item in filtered_data if item.get('status') == status_target]
    return render_template('laporan.html', list_permohonan=filtered_data)

@app.route('/export-laporan-drive')
def export_laporan_drive():
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    generate_and_upload_laporan_pdf(get_all_permohonan())
    return redirect(url_for('laporan_page'))

@app.route('/pengguna')
def pengguna_page():
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    return render_template('pengguna.html', data_pengguna=get_all_pengguna())

@app.route('/pengaturan')
def pengaturan_page():
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    return render_template('pengaturan.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_page'))

@app.route('/permohonan/<kategori>')
def permohonan_page(kategori):
    if not session.get('user'):
        return redirect(url_for('login_page'))
    kategori_valid = ['sarana', 'prasarana', 'instalasi', 'peralatan']
    if kategori not in kategori_valid:
        return redirect(url_for('dashboard_page'))
    return render_template('permohonan.html', kategori=kategori)

@app.route('/proses-permohonan/<kategori>', methods=['POST'])
def proses_permohonan(kategori):
    if not session.get('user'):
        return redirect(url_for('login_page'))
    file_list = []
    drive_ids = []
    def handle_file_upload(input_name):
        files = request.files.getlist(input_name)
        for file in files:
            if file and file.filename != '':
                filename = secure_filename(file.filename)
                save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(save_path)
                file_list.append(filename)
                file.seek(0)
                drive_id = upload_to_drive(file, kategori)
                if drive_id:
                    drive_ids.append(drive_id)
    handle_file_upload('foto')
    handle_file_upload('dokumen')
    permohonan_baru = {
        'pemohon': session.get('nama', 'User'),
        'kategori': kategori,
        'jenis_sarana': request.form.get('jenis_sarana'),
        'jenis_prasarana': request.form.get('jenis_prasarana'),
        'jenis_instalasi': request.form.get('jenis_instalasi'),
        'jenis_peralatan': request.form.get('jenis_peralatan'),
        'tahun_pembuatan': request.form.get('tahun_pembuatan'),
        'kapasitas_orang': request.form.get('kapasitas_orang'),
        'no_lambung': request.form.get('no_lambung'),
        'no_polisi': request.form.get('no_polisi'),
        'merk': request.form.get('merk'),
        'tipe': request.form.get('tipe'),
        'nomer_mesin': request.form.get('nomer_mesin'),
        'nomer_rangka': request.form.get('nomer_rangka'),
        'nomer_stnk': request.form.get('nomer_stnk'),
        'instansi': request.form.get('instansi'),
        'nama_prasarana': request.form.get('nama_prasarana'),
        'tahun_konstruksi': request.form.get('tahun_konstruksi'),
        'lokasi': request.form.get('lokasi'),
        'koordinat': request.form.get('koordinat'),
        'tahun': request.form.get('tahun'),
        'kapasitas': request.form.get('kapasitas'),
        'no_sertifikat': request.form.get('no_sertifikat'),
        'tgl_berlaku': request.form.get('tgl_berlaku'),
        'tgl_berakhir': request.form.get('tgl_berakhir'),
        'catatan': request.form.get('catatan'),
        'status': 'Pending',
        'foto': file_list,
        'foto_drive_ids': drive_ids
    }
    add_permohonan(permohonan_baru)
    return redirect(url_for('riwayat_page'))

@app.route('/riwayat')
def riwayat_page():
    if not session.get('user'):
        return redirect(url_for('login_page'))
    current_user = session.get('nama', 'User')
    user_permohonan = get_permohonan_by_pemohon(current_user)
    return render_template('riwayat.html', list_riwayat=user_permohonan)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=True)