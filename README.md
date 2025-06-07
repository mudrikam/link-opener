# 🔗 Link Opener - Ekstrak & Buka Link dari Dokumen

Aplikasi desktop yang memudahkan kamu untuk mengekstrak dan membuka semua link dari berbagai jenis file dokumen langsung ke Google Chrome. Sangat berguna untuk batch opening link dari file yang berisi banyak URL.

## ✨ Fitur Utama

### 📂 Mendukung Banyak Format File
- **Text Files**: TXT
- **Microsoft Office**: DOC, DOCX, XLS, XLSX, PPT, PPTX  
- **OpenDocument**: ODT (Writer), ODS (Calc), ODP (Impress)
- **Other Formats**: PDF, CSV, RTF

### 🖱️ Interface yang User-Friendly
- **Drag & Drop**: Seret file langsung ke aplikasi
- **File Browser**: Pilih file menggunakan dialog
- **Preview Links**: Lihat semua link yang ditemukan sebelum dibuka
- **Batch Open**: Buka semua link sekaligus di Chrome incognito
- **Individual Open**: Double-click untuk buka link satu per satu
- **Export Links**: Simpan daftar link ke file TXT

### 🚀 Kontrol Chrome yang Canggih
- Otomatis buka Chrome dalam mode incognito
- Track semua tab yang dibuka dari aplikasi
- Tutup semua tab yang dibuka dengan satu klik
- Progress tracking real-time saat membuka link

## 📋 Persyaratan Sistem

- **OS**: Windows 10/11
- **Browser**: Google Chrome (harus ter-install)
- **Python**: 3.8 atau lebih baru
- **RAM**: Minimal 4GB (8GB recommended untuk banyak link)
- **Storage**: 100MB free space

## 🛠️ Cara Install (Step by Step)

### Langkah 1: Install Python

1. **Download Python**:
   - Buka browser dan kunjungi: https://www.python.org/downloads/
   - Klik tombol "Download Python 3.x.x" (versi terbaru)
   - Tunggu download selesai

2. **Install Python**:
   - Double-click file installer yang sudah didownload
   - ⚠️ **PENTING**: Centang "Add Python to PATH" di bagian bawah
   - Klik "Install Now"
   - Tunggu instalasi selesai
   - Klik "Close"

3. **Verifikasi Installation**:
   - Tekan `Windows + R`
   - Ketik `cmd` dan tekan Enter
   - Di Command Prompt, ketik: `python --version`
   - Jika muncul "Python 3.x.x", instalasi berhasil

### Langkah 2: Download Link Opener

1. **Download aplikasi**:
   - Download semua file aplikasi ke folder, misalnya: `C:\LinkOpener\`
   - Pastikan file-file ini ada:
     - `main.py`
     - `requirements.txt`
     - `chromedriver.exe`
     - `link_opener.ico`
     - `run.bat`

### Langkah 3: Install Dependencies

1. **Buka Command Prompt di folder aplikasi**:
   - Buka folder `C:\LinkOpener\` di File Explorer
   - Klik di address bar, ketik `cmd` dan tekan Enter
   - Command Prompt akan terbuka di folder tersebut

2. **Install packages yang dibutuhkan**:
   ```bash
   pip install -r requirements.txt
   ```
   - Tunggu proses download dan instalasi selesai (beberapa menit)
   - Jika ada error, coba: `python -m pip install -r requirements.txt`

### Langkah 4: Verifikasi Chrome

1. **Pastikan Google Chrome ter-install**:
   - Jika belum punya Chrome, download dari: https://www.google.com/chrome/
   - Install Chrome dengan setting default

2. **Test chromedriver**:
   - File `chromedriver.exe` sudah disertakan dalam aplikasi
   - Pastikan file ini ada di folder yang sama dengan `main.py`

## 🚀 Cara Menjalankan Aplikasi

### Metode 1: Double-click (Termudah)
- Double-click file `run.bat`
- Aplikasi akan terbuka otomatis

### Metode 2: Command Line
1. Buka Command Prompt di folder aplikasi
2. Jalankan perintah:
   ```bash
   python main.py
   ```

## 📖 Cara Menggunakan

### 1. Pilih File
**Metode A - Drag & Drop:**
- Seret file dokumen ke area "Drag & Drop" di aplikasi
- File akan otomatis diproses

**Metode B - File Browser:**
- Klik tombol "Pilih File"
- Pilih file dari dialog yang muncul

### 2. Preview Links
- Setelah file diproses, semua link akan muncul di tabel
- Scroll untuk melihat semua link yang ditemukan
- Klik kanan pada link untuk menu context (Copy/Open)

### 3. Buka Links
**Buka Semua Sekaligus:**
- Klik "Buka Chrome" untuk membuka semua link di Chrome incognito
- Progress bar akan menunjukkan kemajuan

**Buka Satu per Satu:**
- Double-click pada link di tabel untuk membuka individual

### 4. Kelola Tab Chrome
- Klik "Tutup Tab" untuk menutup semua tab yang dibuka dari aplikasi
- Aplikasi akan track tab mana saja yang dibukanya

### 5. Export Links
- Klik "Export Links" untuk menyimpan daftar link ke file TXT
- File akan disimpan dengan nama `[namafile]_links.txt`

## 📄 Format File yang Didukung

### Text Files
- **TXT**: File teks biasa dengan encoding UTF-8, Latin-1, atau CP1252

### Microsoft Office
- **DOC**: Word 97-2003 (ekstraksi basic)
- **DOCX**: Word 2007+ dengan hyperlink detection
- **XLS**: Excel 97-2003 dengan hyperlink detection  
- **XLSX**: Excel 2007+ dengan hyperlink detection
- **PPT**: PowerPoint 97-2003 (ekstraksi basic)
- **PPTX**: PowerPoint 2007+ dengan hyperlink detection

### OpenDocument
- **ODT**: OpenDocument Text (LibreOffice Writer)
- **ODS**: OpenDocument Spreadsheet (LibreOffice Calc)
- **ODP**: OpenDocument Presentation (LibreOffice Impress)

### Other Formats
- **PDF**: Portable Document Format dengan link annotation
- **CSV**: Comma-Separated Values dengan berbagai encoding
- **RTF**: Rich Text Format

## 🔍 Contoh File Input

File bisa berisi campuran teks dan link dalam format apapun:

```
Laporan Meeting Project Alpha
=============================

Website referensi:
- Documentation: https://docs.python.org/3/
- GitHub Repository: https://github.com/username/project
- Live Demo: https://project-demo.herokuapp.com

Catatan tambahan:
Untuk informasi lebih lanjut, kunjungi www.example.com
atau hubungi team di https://slack.company.com/channels/alpha

File ini juga tersedia di: https://drive.google.com/file/d/abc123
```

## ⚠️ Troubleshooting

### "chromedriver.exe tidak ditemukan"
- Pastikan file `chromedriver.exe` ada di folder yang sama dengan `main.py`
- Re-download aplikasi jika file hilang

### "Module tidak ditemukan"
- Jalankan ulang: `pip install -r requirements.txt`
- Pastikan menggunakan Python yang benar jika punya multiple versions

### "Chrome tidak bisa dibuka"
- Pastikan Google Chrome ter-install
- Restart aplikasi jika Chrome crash
- Check apakah Chrome sudah update ke versi terbaru

### "File tidak bisa dibaca"
- Pastikan file tidak sedang dibuka di aplikasi lain
- Coba convert file ke format yang lebih umum (misal DOC ke DOCX)
- Check apakah file tidak corrupt

### Link tidak terdeteksi
- Pastikan link dalam format yang benar (http:// atau https://)
- Link tidak boleh ada spasi di tengah
- Coba copy-paste link manual untuk test

## 🤝 Dukungan

Jika mengalami masalah:
1. Pastikan semua langkah instalasi sudah diikuti dengan benar
2. Restart komputer setelah install Python
3. Coba jalankan dengan file `contoh_links.txt` yang disertakan untuk test
4. Check apakah antivirus memblok aplikasi

## 📁 File yang Disertakan

- `main.py` - Aplikasi utama
- `requirements.txt` - Daftar library yang dibutuhkan  
- `chromedriver.exe` - Driver untuk kontrol Chrome
- `link_opener.ico` - Icon aplikasi
- `run.bat` - Script untuk menjalankan aplikasi
- `contoh_links.txt` - File contoh untuk testing
- `README.md` - Dokumentasi lengkap ini
