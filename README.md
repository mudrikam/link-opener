# Link Opener - Aplikasi PySide6

Aplikasi sederhana untuk membuka link dari file TXT ke Google Chrome.

## Fitur
- GUI sederhana dengan tombol pilih file dan progress bar
- Otomatis mendeteksi link HTTP/HTTPS dari file txt yang berisi campuran teks
- Membuka semua link yang ditemukan di Google Chrome
- Progress tracking saat membuka link
- Preview link yang ditemukan sebelum dibuka

## Cara Install

1. Install Python 3.8+ jika belum ada
2. Install dependency:
```bash
pip install -r requirements.txt
```

## Cara Pakai

1. Jalankan aplikasi:
```bash
python main.py
```

2. Klik tombol "📁 Pilih File TXT"
3. Pilih file .txt yang berisi campuran teks dan link
4. Aplikasi akan menampilkan preview link yang ditemukan
5. Klik "🚀 Buka Semua Link di Chrome" untuk membuka semua link

## Format File TXT

File txt bisa berisi campuran teks dan link. Aplikasi akan otomatis mendeteksi:
- Link yang dimulai dengan `http://` atau `https://`
- Link yang tidak terputus (tanpa spasi di tengah)

Contoh isi file:
```
Ini teks biasa.
https://www.google.com
Teks lain.
https://www.youtube.com/watch?v=abc123
Email: test@example.com (ini tidak akan dibuka)
http://github.com
```

## Persyaratan Sistem

- Windows (aplikasi ini dibuat khusus untuk Windows)
- Google Chrome ter-install
- Python 3.8+
- PySide6

## File yang Disertakan

- `main.py` - Aplikasi utama
- `requirements.txt` - Dependency Python
- `contoh_links.txt` - Contoh file untuk testing
- `README.md` - Dokumentasi ini
