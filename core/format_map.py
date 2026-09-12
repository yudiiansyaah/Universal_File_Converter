"""
format_map.py
Peta kategori dan format konversi yang didukung aplikasi.
Struktur ini mengikuti tabel referensi "Daftar Format File yang Dapat Dikonversi"
yang diberikan user, dengan penyesuaian pada format yang secara teknis
dapat diimplementasikan secara otomatis (tanpa intervensi manual).
"""

# Setiap entri: ekstensi sumber -> daftar ekstensi tujuan yang valid
DOCUMENT = {
    "docx": ["pdf", "txt", "md", "rtf", "html", "odt"],
    "doc":  ["docx", "pdf", "txt", "rtf"],
    "odt":  ["docx", "pdf", "txt", "html"],
    "rtf":  ["docx", "pdf", "txt", "md"],
    "txt":  ["pdf", "docx", "md", "rtf", "html"],
    "md":   ["pdf", "docx", "html", "txt", "rtf"],
    "html": ["pdf", "docx", "txt", "md"],
    "pdf":  ["docx", "txt"],  # md/html dari pdf kualitasnya rendah, disediakan lewat docx->md manual bila perlu
}

SPREADSHEET = {
    "xlsx": ["csv", "pdf", "ods", "json", "html"],
    "xls":  ["xlsx", "csv", "pdf"],
    "csv":  ["xlsx", "pdf", "ods", "txt", "json"],
    "ods":  ["xlsx", "csv", "pdf"],
    "tsv":  ["csv", "xlsx", "txt"],
    # json/xml/yaml bisa berupa data tabular (list of records) ATAU data
    # terstruktur bebas (nested). Target mencakup keduanya; dispatcher di
    # converters.py yang menentukan jalur pemrosesan berdasarkan target.
    "json": ["csv", "xlsx", "txt", "md", "yaml", "xml"],
    "xml":  ["csv", "xlsx", "json", "txt", "yaml"],
    "yaml": ["json", "txt", "md", "xml", "csv"],
    "yml":  ["json", "txt", "md", "xml", "csv"],
}

PRESENTATION = {
    "pptx": ["pdf", "ppt", "odp"],
    "ppt":  ["pptx", "pdf"],
    "odp":  ["pptx", "pdf"],
}

IMAGE = {
    "png":  ["jpg", "jpeg", "webp", "bmp", "tiff", "pdf"],
    "jpg":  ["png", "webp", "bmp", "tiff", "pdf"],
    "jpeg": ["png", "webp", "bmp", "tiff", "pdf"],
    "webp": ["png", "jpg", "jpeg"],
    "gif":  ["png", "jpg", "webp"],
    "bmp":  ["png", "jpg", "webp", "tiff"],
    "tiff": ["png", "jpg", "jpeg", "webp"],
    "tif":  ["png", "jpg", "jpeg", "webp"],
    "svg":  ["png", "jpg", "webp", "pdf"],
}

AUDIO = {
    "mp3":  ["wav", "m4a", "flac", "ogg", "aac", "opus"],
    "wav":  ["mp3", "m4a", "flac", "ogg"],
    "m4a":  ["mp3", "wav", "flac", "ogg"],
    "flac": ["mp3", "wav", "m4a", "ogg"],
    "ogg":  ["mp3", "wav", "m4a"],
    "aac":  ["mp3", "wav", "m4a"],
    "opus": ["mp3", "wav", "ogg"],
}

VIDEO = {
    "mp4":  ["mkv", "avi", "mov", "webm"],
    "mkv":  ["mp4", "avi", "mov", "webm"],
    "avi":  ["mp4", "mkv", "mov", "webm"],
    "mov":  ["mp4", "mkv", "avi", "webm"],
    "webm": ["mp4", "mkv", "mov"],
    "flv":  ["mp4", "mkv", "avi"],
    "wmv":  ["mp4", "mkv", "avi"],
    "mpeg": ["mp4", "mkv", "avi"],
    "mpg":  ["mp4", "mkv", "avi"],
    "m4v":  ["mp4", "mkv", "mov"],
}

ARCHIVE = {
    "zip":     ["tar", "tar.gz", "7z"],
    "tar":     ["zip", "tar.gz", "tar.xz"],
    "tar.gz":  ["zip", "tar", "7z"],
    "tgz":     ["zip", "tar", "7z"],
    "tar.bz2": ["zip", "tar.gz", "tar.xz"],
    "tar.xz":  ["zip", "tar.gz", "tar.bz2"],
    "7z":      ["zip", "tar.gz"],
    "rar":     ["zip", "tar.gz"],  # ekstraksi rar butuh 'unrar'/'unar' terpasang
}

# Format data terstruktur (bagian dari kategori "Kode & Konfigurasi" yang
# BENAR-BENAR bisa dikonversi otomatis tanpa memahami semantik source code).
# json/xml/yaml/yml SENGAJA tidak didaftarkan ulang sebagai key di sini
# (sudah ada di SPREADSHEET) agar tiap ekstensi sumber hanya masuk satu
# kategori dan tidak ambigu.
CODE_DATA = {
    "toml": ["json", "yaml"],
    "ini":  ["json", "yaml", "txt"],
    "conf": ["txt", "json"],
}

# Bahasa/skrip pemrograman: transformasi source code TIDAK dilakukan otomatis
# (butuh compiler/transpiler khusus: tsc, dsb). Ditampilkan di UI sebagai info,
# bukan sebagai target konversi yang bisa dieksekusi.
CODE_UNSUPPORTED_NOTE = (
    "Konversi source code (py, js, ts, java, c/cpp, cs, php, sh/bash, ps1, sql, css) "
    "memerlukan transpiler/compiler khusus per bahasa dan tidak dilakukan otomatis "
    "oleh aplikasi ini untuk menghindari hasil yang salah secara semantik."
)

ALL_CATEGORIES = {
    "Dokumen & Teks": DOCUMENT,
    "Spreadsheet & Data": SPREADSHEET,
    "Presentasi": PRESENTATION,
    "Gambar": IMAGE,
    "Audio": AUDIO,
    "Video": VIDEO,
    "Arsip & Kompresi": ARCHIVE,
    "Data Terstruktur (Kode & Konfigurasi)": CODE_DATA,
}


def find_targets(ext: str):
    """Cari daftar format tujuan yang valid untuk sebuah ekstensi sumber."""
    ext = ext.lower().lstrip(".")
    for category in ALL_CATEGORIES.values():
        if ext in category:
            return category[ext]
    return []


def find_category(ext: str):
    ext = ext.lower().lstrip(".")
    for name, category in ALL_CATEGORIES.items():
        if ext in category:
            return name
    return None
