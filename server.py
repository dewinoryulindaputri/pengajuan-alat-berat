from flask import Flask, render_template, request, redirect, url_for, session
import os
import io
import pandas as pd
import openpyxl
from flask import send_file
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
# --- MODIFIKASI: Import untuk OAuth ---
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle

app = Flask(__name__)
app.secret_key = 'kunci_rahasia_spip'

# --- KONFIGURASI GOOGLE DRIVE ---
SERVICE_ACCOUNT_FILE = 'client_secret.json'
SCOPES = ['https://www.googleapis.com/auth/drive.file']
GOOGLE_DRIVE_FOLDER_ID = '1H8crWdaeqPcUHrlpqaqrs4F2WoY89nzg'

# --- MODIFIKASI: Fungsi OAuth diperbarui (aman jika token.pickle belum ada) ---
def get_drive_service():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(SERVICE_ACCOUNT_FILE, SCOPES)
            # Menggunakan port=0 agar otomatis menyesuaikan dengan tipe Desktop App
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
            
    return build('drive', 'v3', credentials=creds)

def upload_to_drive(file_storage):
    try:
        service = get_drive_service()
        file_content = file_storage.read()
        media_body = MediaIoBaseUpload(io.BytesIO(file_content), mimetype=file_storage.content_type, resumable=True)
        file_metadata = {'name': file_storage.filename, 'parents': [GOOGLE_DRIVE_FOLDER_ID]}
        file = service.files().create(body=file_metadata, media_body=media_body, fields='id').execute()
        return file.get('id')
    except Exception as e:
        print(f"Gagal upload ke Drive: {e}")
        return None

# Konfigurasi folder untuk menyimpan hasil upload foto
UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# DATABASE SEMENTARA
data_permohonan = []
data_sarana = [
    {
        'id': 1,
        'jenis_sarana': 'Dump Truck',
        'no_lambung': 'DT-001',
        'merk': 'Scania',
        'tipe': 'P460',
        'instansi': 'Departemen Tambang',
        'status': 'Aktif'
    }
]

data_pengguna = [
    {'id': 1, 'nama': 'Super Administrator', 'email': 'admin@perizinan.com', 'role': 'Administrator'},
    {'id': 2, 'nama': 'Petugas Verifikasi', 'email': 'petugas@perizinan.com', 'role': 'Petugas'}
]

@app.route('/')
def login_page():
    return render_template('login.html')

@app.route('/login', methods=['POST'])
def do_login():
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '').strip()
    session.clear() 
    
    if username == 'admin' and password == '12345':
        session['user'] = 'admin'
        session['nama'] = 'Administrator'
    else:
        nama_user = username if username else 'User'
        session['user'] = 'user'
        session['nama'] = nama_user
        sudah_ada = any(p['nama'] == nama_user for p in data_pengguna)
        if not sudah_ada:
            data_pengguna.append({
                'id': len(data_pengguna) + 1,
                'nama': nama_user,
                'email': f"{nama_user.lower().replace(' ', '')}@perizinan.com",
                'role': 'User / Pemohon'
            })
    return redirect(url_for('dashboard_page'))

@app.route('/dashboard')
def dashboard_page():
    role = session.get('user')
    if role == 'admin':
        return redirect(url_for('admin_page'))
    elif role == 'user':
        return render_template('index.html')
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
    
    foto_filenames = []
    drive_ids = []
    files = request.files.getlist('foto')
    
    for file in files:
        if file and file.filename != '':
            filename = file.filename
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            foto_filenames.append(filename)
            file.seek(0)
            drive_id = upload_to_drive(file)
            if drive_id:
                drive_ids.append(drive_id)
    
    permohonan_baru = {
        'id': len(data_permohonan) + 1,
        'kode_unik': f"REQ-{len(data_permohonan) + 1:03d}",
        'pemohon': session.get('nama', 'User'),
        'kategori': kategori,
        'jenis_prasarana': request.form.get('jenis_prasarana'),
        'nama_prasarana': request.form.get('nama_prasarana'),
        'lokasi': request.form.get('lokasi'),
        'koordinat': request.form.get('koordinat'),
        'jenis_sarana': request.form.get('jenis_sarana'),
        'tahun': request.form.get('tahun'),
        'no_lambung': request.form.get('no_lambung'),
        'aktivitas': request.form.get('aktivitas'),
        'no_polisi': request.form.get('no_polisi'),
        'instansi': request.form.get('instansi'),
        'merk': request.form.get('merk'),
        'kapasitas': request.form.get('kapasitas'),
        'tipe': request.form.get('type'),
        'nomer_mesin': request.form.get('nomer_mesin'),
        'nomer_rangka': request.form.get('nomer_rangka'),
        'nomer_stnk': request.form.get('nomer_stnk'),
        'perusahaan_user': request.form.get('perusahaan_user'),
        'nomer_wa': request.form.get('nomer_wa'),
        'catatan': request.form.get('catatan'),
        'status': 'Pending',
        'foto': foto_filenames,
        'foto_drive_ids': drive_ids
    }
    
    data_permohonan.append(permohonan_baru)
    return redirect(url_for('riwayat_page'))

@app.route('/riwayat')
def riwayat_page():
    if not session.get('user'):
        return redirect(url_for('login_page'))
    current_user = session.get('nama', 'User')
    user_permohonan = [item for item in data_permohonan if item['pemohon'] == current_user]
    return render_template('riwayat.html', list_riwayat=user_permohonan)

@app.route('/detail-permohonan/<int:id>')
def detail_permohonan_page(id):
    if not session.get('user'):
        return redirect(url_for('login_page'))
    permohonan = next((item for item in data_permohonan if item['id'] == id), None)
    if not permohonan:
        return redirect(url_for('dashboard_page'))
    return render_template('detail_permohonan.html', permohonan=permohonan)

@app.route('/admin')
def admin_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('admin.html')

@app.route('/permohonan-masuk')
def permohonan_masuk_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('permohonan_masuk.html', list_permohonan=data_permohonan)

@app.route('/setujui/<int:id>')
def setujui_permohonan(id):
    if session.get('user') != 'admin': return redirect(url_for('dashboard_page'))
    for item in data_permohonan:
        if item['id'] == id:
            item['status'] = 'Disetujui'
            break
    return redirect(url_for('permohonan_masuk_page'))

@app.route('/tolak/<int:id>')
def tolak_permohonan(id):
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    for item in data_permohonan:
        if item['id'] == id:
            item['status'] = 'Ditolak'
            break
    return redirect(url_for('permohonan_masuk_page'))

@app.route('/pengaturan')
def pengaturan_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('pengaturan.html')

@app.route('/master-sarana')
def master_sarana_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('master_sarana.html', data_sarana=data_sarana)

@app.route('/master-prasarana')
def master_prasarana_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('master_prasarana.html')

@app.route('/master-instalasi')
def master_instalasi_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('master_instalasi.html')

@app.route('/master-peralatan')
def master_peralatan_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('master_peralatan.html')

@app.route('/tambah-sarana')
def tambah_sarana_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('tambah_sarana.html')

@app.route('/edit-sarana')
def edit_sarana_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('edit_sarana.html')

@app.route('/pengguna')
def pengguna_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('pengguna.html', data_pengguna=data_pengguna)

@app.route('/tambah-pengguna')
def tambah_pengguna_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('tambah_pengguna.html')

@app.route('/manajemen-pengguna')
def manajemen_pengguna_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('manajemen_pengguna.html')

@app.route('/laporan')
def laporan_page():
    if session.get('user') != 'admin':
        return redirect(url_for('dashboard_page'))
    return render_template('laporan.html', list_permohonan=data_permohonan)

@app.route('/export-excel')
def export_excel():
    if session.get('user') != 'admin':
        return redirect(url_for('login_page'))
    
    template_path = 'FORM ISIAN SPIP.xlsx'
    if not os.path.exists(template_path):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for kat in ['SARANA', 'PRASARANA', 'INSTALASI', 'PERALATAN']:
            wb.create_sheet(title=kat)
        wb.save(template_path)
        
    try:
        wb = openpyxl.load_workbook(template_path)
        sheet_mapping = {
            'sarana': 'SARANA',
            'prasarana': 'PRASARANA',
            'instalasi': 'INSTALASI',
            'peralatan': 'PERALATAN'
        }
        
        for item in data_permohonan:
            kat = item.get('kategori', '').strip().lower()
            sheet_name = sheet_mapping.get(kat)
            
            if sheet_name and sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                if kat == 'sarana':
                    row_data = [item.get('jenis_sarana', ''), item.get('no_lambung', ''), item.get('no_polisi', ''), item.get('merk', ''), item.get('tipe', ''), item.get('tahun', ''), item.get('nomer_mesin', ''), item.get('nomer_rangka', ''), item.get('nomer_stnk', ''), item.get('kapasitas', ''), item.get('perusahaan_user', ''), item.get('instansi', ''), item.get('pemohon', ''), item.get('nomer_wa', ''), ', '.join(item.get('foto', []))]
                elif kat == 'prasarana':
                    row_data = [item.get('jenis_prasarana', ''), item.get('nama_prasarana', ''), item.get('tahun', ''), item.get('lokasi', ''), item.get('koordinat', ''), item.get('kapasitas', ''), ', '.join(item.get('foto', []))]
                elif kat == 'instalasi':
                    row_data = [item.get('jenis_instalasi', ''), item.get('lokasi', ''), item.get('tahun', ''), item.get('kapasitas', ''), item.get('no_sertifikat', ''), item.get('tgl_berlaku', ''), item.get('tgl_berakhir', ''), ', '.join(item.get('foto', []))]
                elif kat == 'peralatan':
                    row_data = [item.get('jenis_peralatan', ''), item.get('no_lambung', ''), item.get('merk', ''), item.get('tipe', ''), item.get('nomer_mesin', ''), item.get('nomer_rangka', ''), item.get('tahun', ''), item.get('kapasitas', ''), item.get('no_sertifikat', ''), item.get('tgl_berlaku', ''), item.get('tgl_berakhir', ''), ', '.join(item.get('foto', []))]
                else: continue
                ws.append(row_data)
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='FORM_ISIAN_SPIP_Terisi.xlsx')
    except Exception as e:
        print(f"Gagal memproses file Excel: {e}")
        return "Terjadi kesalahan saat memproses file Excel.", 500

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_page'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=True)