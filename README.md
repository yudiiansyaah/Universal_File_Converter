# Universal File Converter (Windows 11)

Aplikasi desktop untuk mengonversi file secara **batch** antar berbagai format,
mengikuti daftar referensi format yang kamu berikan: Dokumen, Spreadsheet,
Presentasi, Gambar, Audio, Video, Arsip, dan Data Terstruktur (JSON/YAML/XML/TOML/INI).

Dibangun dengan **Python + PySide6 (Qt)**. Bisa dijalankan langsung dengan Python,
atau di-build jadi satu file `.exe` mandiri dengan PyInstaller.

---

## 1. Menjalankan langsung (mode developer)

Syarat: **Python 3.10+** terpasang di Windows (centang "Add to PATH" saat instal).

```bat
cd converter_app
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## 2. Build jadi file .exe (untuk didistribusikan)

Setelah langkah di atas (venv aktif, dependencies terpasang), jalankan:

```bat
build.bat
```

Hasilnya ada di `dist\UniversalFileConverter.exe` — file tunggal yang bisa
di-double click tanpa perlu Python terpasang di komputer lain.

> PyInstaller harus dijalankan **di Windows** untuk menghasilkan `.exe` Windows
> (tidak bisa cross-compile dari Linux/Mac). Kalau kamu develop di Linux/WSL,
> pindahkan folder ini ke Windows dulu sebelum `build.bat`.

## 3. Bikin installer "Next-Next-Install" (untuk user awam)

`build.bat` otomatis akan mencoba bikin installer kalau **Inno Setup** sudah
terpasang. Kalau belum:

1. Download & install Inno Setup gratis di https://jrsoftware.org/isinfo.php
   (pilih versi terbaru, cukup Next-Next-Install seperti biasa).
2. Jalankan ulang `build.bat`, ATAU langsung:
   ```bat
   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
   ```
3. Hasilnya: `installer_output\UniversalFileConverterSetup.exe`

**Inilah file yang kamu bagikan ke user awam.** Mereka tinggal double-click,
klik Next beberapa kali, Install, selesai — muncul shortcut di Start Menu
(dan Desktop kalau dicentang saat instalasi). Tidak perlu Python, tidak perlu
command line sama sekali di sisi user.

> Konfigurasi installer (nama app, versi, ikon, shortcut) ada di `installer.iss`
> — edit sesuai kebutuhan sebelum compile ulang.

---

## 4. Dependensi eksternal (WAJIB untuk sebagian fitur)

Aplikasi ini murni Python untuk: **Gambar, Arsip, Data Terstruktur (JSON/YAML/
XML/TOML/INI), dan sebagian Spreadsheet (csv/xlsx/txt/json)**. Tidak perlu
instalasi tambahan untuk kategori-kategori itu.

Tapi untuk kategori berikut, aplikasi memanggil tool eksternal (industry-standard,
supaya hasil konversi akurat, bukan reinvent-the-wheel yang berisiko rusak):

### a) FFmpeg (Audio & Video) — wajib
1. Download dari https://www.gyan.dev/ffmpeg/builds/ (pilih "release full").
2. Ekstrak, lalu tambahkan folder `bin` (isinya `ffmpeg.exe`) ke **Environment
   Variables > PATH**.
3. Cek dengan buka Command Prompt baru, ketik `ffmpeg -version`.

### b) Dokumen & Presentasi — pilih salah satu, TIDAK wajib LibreOffice
Aplikasi otomatis memilih salah satu dari 3 jalur berikut, sesuai apa yang
terdeteksi di komputer (urutan prioritas dari atas ke bawah):

1. **LibreOffice** (kalau terpasang) — kualitas terbaik, mendukung semua
   kombinasi format termasuk Presentasi (pptx/ppt/odp).
   Download di https://www.libreoffice.org/download/. Terdeteksi otomatis
   di `C:\Program Files\LibreOffice\program\`.
2. **Microsoft Word** (kalau LibreOffice tidak ada, tapi Word terpasang) —
   dipakai otomatis lewat automation, hasil setara LibreOffice untuk
   kategori Dokumen (docx/doc/rtf/txt/html/pdf). Butuh `pip install pywin32`
   (sudah ada di `requirements.txt`).
3. **Mode ringan (tanpa Word/LibreOffice sama sekali)** — kalau keduanya
   tidak ada, aplikasi tetap bisa convert **file .docx** ke `.txt`, `.md`,
   `.html`, dan `.pdf` pakai `python-docx` + `reportlab` murni Python
   (juga sudah ada di `requirements.txt`). **Keterbatasan:** hanya
   mengekstrak teks paragraf — formatting asli (bold, tabel, gambar,
   layout kolom, dst) TIDAK dipertahankan, dan hanya mendukung sumber
   `.docx` (bukan `.doc`/`.odt`/`.rtf`). Presentasi (pptx) tetap butuh
   LibreOffice, tidak ada mode ringan untuk itu.

> **Kalau file kamu dari Google Docs:** buka file-nya di Google Docs, lalu
> **File > Download > Microsoft Word (.docx)** dulu sebelum dimasukkan ke
> aplikasi ini — Google Docs online tidak bisa diakses langsung oleh app
> desktop.

Aplikasi menampilkan **banner info** di bagian atas kalau LibreOffice/Word
sama-sama tidak terdeteksi (bukan error, cuma pemberitahuan bahwa hasil
Dokumen akan pakai mode ringan). Untuk Presentasi, error baru muncul saat
kamu benar-benar mencoba convert file pptx tanpa LibreOffice.

---

## 5. Cara pakai aplikasi

1. Klik **"+ Tambah File"**, pilih satu atau banyak file sekaligus.
2. Aplikasi otomatis mendeteksi format sumber dan menampilkan pilihan format
   tujuan yang valid di dropdown kolom "Konversi ke" (per baris/file, bisa beda-beda).
3. (Opsional) Pilih **Folder Output** — kalau dikosongkan, hasil disimpan di
   folder yang sama dengan file sumber masing-masing.
4. Klik **"Konversi Semua"**. Progress & status per file terlihat di tabel,
   dan log detail (termasuk pesan error) muncul di panel bawah.

---

## 6. Cakupan format (mengikuti dokumen referensi)

| Kategori | Metode |
|---|---|
| Dokumen (docx, doc, odt, rtf, txt, md, html, pdf) | LibreOffice + fallback ringan Python |
| Spreadsheet (xlsx, xls, csv, ods, tsv, json, xml, yaml) | pandas/openpyxl + LibreOffice untuk render PDF/ODS |
| Presentasi (pptx, ppt, odp, pdf) | LibreOffice |
| Gambar (png, jpg, webp, bmp, tiff, gif, svg, pdf) | Pillow + cairosvg (untuk SVG) |
| Audio (mp3, wav, m4a, flac, ogg, aac, opus) | FFmpeg |
| Video (mp4, mkv, avi, mov, webm, flv, wmv, mpeg, m4v) | FFmpeg |
| Arsip (zip, tar, tar.gz, tar.bz2, tar.xz, 7z, rar*) | zipfile/tarfile/py7zr (*ekstrak rar butuh tool `unrar` tambahan) |
| Data Terstruktur (json, yaml, xml, toml, ini, conf) | Python stdlib + pyyaml/toml |

### Catatan yang SENGAJA tidak diimplementasikan otomatis
Konversi **source code antar bahasa pemrograman** (py, js, ts, java, c/cpp, cs,
php, sh/bash, ps1, sql, css) **tidak** dilakukan otomatis. Ini butuh
compiler/transpiler khusus per bahasa (mis. `tsc` untuk TypeScript), dan
konversi "asal jalan" berisiko menghasilkan kode yang salah secara semantik —
lebih baik tidak menyediakan fitur ini daripada memberi hasil yang menyesatkan.
File konfigurasi berisi password/token/API key juga sebaiknya tidak dikonversi
sembarangan karena isinya sensitif (sesuai catatan di dokumen referensimu).

---

## 7. Struktur proyek

```
converter_app/
├── main.py                 # GUI utama (PySide6)
├── core/
│   ├── format_map.py       # Peta format sumber -> tujuan per kategori
│   ├── converters.py       # Logika konversi tiap kategori
│   └── dependencies.py     # Deteksi FFmpeg/LibreOffice
├── assets/
│   └── icon.ico             # Ikon aplikasi & installer
├── requirements.txt
├── build.bat                 # Build ke .exe (PyInstaller) + installer (Inno Setup)
├── installer.iss             # Script Inno Setup
└── README.md
```
