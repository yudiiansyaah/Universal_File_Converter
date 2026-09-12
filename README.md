# Universal File Converter (Windows 11)

## 📥 Download

**[⬇️ Download UniversalFileConverter.exe (versi terbaru)](https://github.com/yudiiansyaah/Universal_File_Converter/releases/latest)**

Klik link di atas Cuy → pilih file `UniversalFileConverter.exe` di bagian **Assets** →
tunggu download sampe selesai → double-click untuk menjalankan. Gak perlu lagi install
Python atau apa pun untuk ORANG AWAM.

> ⚠️ Windows mungkin akan menampilkan peringatan kayak **"Windows protected your PC"**
> saat pertama kali dibuka (karena aplikasi ini belum memakai digital
> signature berbayar). Ini normal Cuy buat software gratis/indie — klik
> **"More info" → "Run anyway"** untuk melanjutkan.
> BTW ini aman dari virus yak cuyy, privasi juga terjaga karena aplikasi ini bisa di running OFFLINE tanpa internet.

---

Aplikasi desktop yang bisa dipake untuk mengonversi file secara **batch** antar berbagai format,
ngikutin kebutuhan format yang lu pengen seperti: Dokumen, Spreadsheet,
Presentasi, Gambar, Audio, Video, Arsip, dan Data Terstruktur (JSON/YAML/XML/TOML/INI).

Dibikin pakai Bahasa Pemrograman **Python + Library PySide6 (Qt)**. Bisa dijalankan langsung pakai Python,
atau di-build sendiri jadi satu file `.exe` dengan PyInstaller cuyy.

> 📦 **Catatan untuk yang meng-clone repo ini:** file `.exe` **GAKK IKUT CUYY**
> di source code (ukurannya ~278MB, melebihi batas ukuran file dari si git-nya) —
> download dari link **Releases** di atas, atau build sendiri dari source
> mengikuti step by step di bawah ini Cuy.

---

## 1. Jalanin langsung (mode developer)

Requirement: **Python 3.10+** udah lu install di Windows lo cuy (centang "Add to PATH" saat instal).

```bat
cd converter_app
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## 2. Cara bikin jadi file .exe (untuk didistribusikan / dibagi ke orang AWAM)

After step di atas (venv aktif, dependencies terpasang), running cuy:

```bat
build.bat
```

Outputnya ada di `dist\UniversalFileConverter.exe` — file sendiri yang bisa
di-double click tanpa perlu install Python buat komputer lain (USER AWAM).

> PyInstaller harus dijalankan **di Windows** untuk menghasilkan `.exe` Windows
> (tidak bisa cross-compile dari Linux/Mac). Kalau kamu develop di Linux/WSL,
> pindahkan folder ini ke Windows dulu sebelum `build.bat`.

## 3. Bikin installer "Tinggal Next-Next-Install" (untuk user awam)

`build.bat` otomatis bakalan coba bikin installer kalau **Inno Setup** sudah
ke install. Kalau belum:

1. Download & install Inno Setup gratis-tis-tis-tis di https://jrsoftware.org/isinfo.php
   (pilih versi terbaru atau like new, cukup Next-Next-Install seperti biasa-nya cuy).
2. Jalankan ulang `build.bat`, ATAU langsung aja cuy:
   ```bat
   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
   ```
3. Hasilnya: `installer_output\UniversalFileConverterSetup.exe`

**Inilah file yang lo bagikan ke user awam.** Mereka tinggal double-click,
klik Next beberapa kali, Install, selesai — muncul shortcut di Start Menu
(dan Desktop kalau dicentang saat instalasi). Tidak perlu Python, tidak perlu
command line sama sekali di sisi user cuy, aman dari eror eror pokoknya.

> Konfigurasi installer (nama app, versi, ikon, shortcut) ada di `installer.iss`
> — edit sesuai kebutuhan sebelum compile ulang cuy.

---

## 4. Dependensi eksternal (WAJIB untuk sebagian fitur)

Aplikasi ini murni bahasa pemrograman Python cuy untuk: **Gambar, Arsip, Data Terstruktur (JSON/YAML/
XML/TOML/INI), dan sebagian Spreadsheet (csv/xlsx/txt/json)**. Gak perlu
instalasi tambahan untuk kategori-kategori itu.

Tapi buat kategori berikut, aplikasi buat memanggil tool eksternal-nya (industry-standard,
supaya hasil konversi akurat, bukan reinvent-the-wheel yang berisiko rusak):

### a) FFmpeg (Audio & Video) — wajib
1. Download dari https://www.gyan.dev/ffmpeg/builds/ (pilih "release full").
2. Ekstrak, lalu tambahkan folder `bin` (isinya `ffmpeg.exe`) ke **Environment
   Variables > PATH**.
3. Cek dengan buka Command Prompt baru, ketik `ffmpeg -version`.

### b) Dokumen & Presentasi — pilih salah satu aja cuy, GAK wajib LibreOffice
Aplikasi otomatis memilih salah satu dari 3 jalur berikut, sesuai apa yang
terdeteksi di komputer lo pada cuy (urutan prioritas dari atas ke bawah):

1. **LibreOffice** (kalau terinstall) — kualitas terbaik, mendukung semua
   kombinasi format termasuk Presentasi (pptx/ppt/odp).
   Download di https://www.libreoffice.org/download/. Terdeteksi otomatis
   di `C:\Program Files\LibreOffice\program\`.
2. **Microsoft Word** (kalau LibreOffice gak ada, tapi Word ke install) —
   dipakai otomatis lewat automation, hasil setara LibreOffice untuk
   kategori Dokumen (docx/doc/rtf/txt/html/pdf). Butuh `pip install pywin32`
   (sudah ada di `requirements.txt`).
3. **Mode ringan (tanpa Word/LibreOffice sama sekali)** — kalau keduanya
   gak ada, aplikasi tetap bisa convert **file .docx** ke `.txt`, `.md`,
   `.html`, dan `.pdf` pakai `python-docx` + `reportlab` murni bahasa pemrograman Python cuy
   (juga sudah ada di `requirements.txt`). **Keterbatasan:** hanya
   mengekstrak teks paragraf — formatting asli (bold, tabel, gambar,
   layout kolom, dst) TIDAK dipertahankan, dan hanya mendukung sumber
   `.docx` (bukan `.doc`/`.odt`/`.rtf`). Presentasi (pptx) tetap butuh
   LibreOffice, tidak ada mode ringan untuk itu.

> **Kalau file lo dari Google Docs:** buka file-nya di Google Docs, lalu
> **File > Download > Microsoft Word (.docx)** dulu sebelum dimasukkan ke
> aplikasi ini — Google Docs online tidak bisa diakses langsung oleh app
> desktop cuyy.

Aplikasi akan menampilkan **banner info** di bagian atas kalau LibreOffice/Word
sama-sama tidak terdeteksi (bukan error cuy, cuma di notice aja kalo hasil
Dokumen akan dipakai mode ringan). Untuk Presentasi, error baru muncul saat
lo benar-benar mencoba convert file pptx tanpa LibreOffice.

---

## 5. Cara pakai aplikasi

1. Klik **"+ Tambah File"**, pilih satu atau banyak file sekaligus.
2. Aplikasi otomatis mendeteksi format sumber dan menampilkan pilihan format
   tujuan yang valid di dropdown kolom "Konversi ke" (per baris/file, bisa beda-beda).
3. (Opsional) Pilih **Folder Output** — kalau dikosongkan, hasil disimpan di
   folder yang sama dengan file sumber masing-masing cuy.
4. Klik **"Konversi Semua"**. Progress & status per file terlihat di tabel,
   dan log detail (termasuk pesan error) muncul di panel bawah cuy.

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

### Catatan yang SENGAJA GAK diimplementasikan otomatis
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
