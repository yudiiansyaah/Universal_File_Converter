"""
converters.py
Implementasi konversi untuk tiap kategori file.
Setiap fungsi convert_* menerima (input_path, output_path) dan
mengembalikan True/False, serta boleh melempar Exception dengan pesan jelas.
"""
import csv
import json
import os
import shutil
import subprocess
import tarfile
import zipfile
from pathlib import Path

from . import dependencies

try:
    import yaml
except ImportError:
    yaml = None

try:
    import toml
except ImportError:
    toml = None

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from PIL import Image
except ImportError:
    Image = None

try:
    import py7zr
except ImportError:
    py7zr = None


class ConversionError(Exception):
    pass


def convert_via_word(input_path: str, target_ext: str, output_dir: str, log=print) -> str:
    """Konversi dokumen lewat Microsoft Word (Windows COM automation)."""
    try:
        import win32com.client
    except ImportError:
        raise ConversionError(
            "Butuh pustaka 'pywin32' untuk memakai Microsoft Word. "
            "Jalankan: pip install pywin32"
        )
    os.makedirs(output_dir, exist_ok=True)
    stem = Path(input_path).stem
    output_path = os.path.abspath(os.path.join(output_dir, f"{stem}.{target_ext}"))
    input_abspath = os.path.abspath(input_path)

    # Kode format Word: https://learn.microsoft.com/office/vba/api/word.wdsaveformat
    wd_formats = {
        "pdf": 17, "docx": 16, "doc": 0, "txt": 2, "rtf": 6, "html": 8,
    }
    if target_ext not in wd_formats:
        raise ConversionError(f"Microsoft Word tidak mendukung ekspor ke '.{target_ext}'.")

    log("Membuka Microsoft Word (mode background)...")
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(input_abspath, ReadOnly=True)
        try:
            log(f"Menyimpan sebagai .{target_ext} lewat Word...")
            doc.SaveAs(output_path, FileFormat=wd_formats[target_ext])
        finally:
            doc.Close(False)
    finally:
        word.Quit()

    if not os.path.isfile(output_path):
        raise ConversionError("Word selesai tapi file output tidak ditemukan.")
    return output_path


def convert_docx_lightweight(input_path: str, target_ext: str, output_path: str, log=print):
    """Fallback murni Python tanpa Word/LibreOffice: ekstrak teks polos dari .docx.
    Layout/format asli TIDAK dipertahankan - hanya isi teksnya."""
    try:
        import docx as docx_lib
    except ImportError:
        raise ConversionError(
            "Butuh pustaka 'python-docx' untuk mode ringan ini. "
            "Jalankan: pip install python-docx"
        )
    if Path(input_path).suffix.lower() != ".docx":
        raise ConversionError(
            "Mode ringan (tanpa Word/LibreOffice) hanya mendukung file .docx, "
            f"bukan .{Path(input_path).suffix.lstrip('.')}."
        )

    doc = docx_lib.Document(input_path)
    paragraphs = [p.text for p in doc.paragraphs]
    text = "\n".join(paragraphs)

    if target_ext == "txt":
        Path(output_path).write_text(text, encoding="utf-8")
        return output_path

    if target_ext == "md":
        Path(output_path).write_text(text, encoding="utf-8")
        return output_path

    if target_ext == "html":
        html_body = "".join(f"<p>{p}</p>" for p in paragraphs if p.strip())
        Path(output_path).write_text(f"<html><body>{html_body}</body></html>", encoding="utf-8")
        return output_path

    if target_ext == "pdf":
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.pdfgen import canvas
            from reportlab.lib.units import cm
        except ImportError:
            raise ConversionError(
                "Butuh pustaka 'reportlab' untuk membuat PDF tanpa Word/LibreOffice. "
                "Jalankan: pip install reportlab"
            )
        log("Membuat PDF sederhana (teks polos, tanpa format asli)...")
        c = canvas.Canvas(output_path, pagesize=A4)
        width, height = A4
        x, y = 2 * cm, height - 2 * cm
        line_height = 14
        for para in paragraphs:
            for line in _wrap_text(para, 90):
                if y < 2 * cm:
                    c.showPage()
                    y = height - 2 * cm
                c.drawString(x, y, line)
                y -= line_height
            y -= 4  # jarak antar paragraf
        c.save()
        return output_path

    raise ConversionError(
        f"Mode ringan tidak mendukung target '.{target_ext}'. "
        "Install LibreOffice atau Microsoft Word untuk dukungan format lebih lengkap."
    )


def _wrap_text(text: str, width: int):
    import textwrap
    if not text.strip():
        return [""]
    return textwrap.wrap(text, width=width) or [""]


# ---------------------------------------------------------------------------
# DOKUMEN & PRESENTASI (via LibreOffice headless)
# ---------------------------------------------------------------------------
def convert_via_libreoffice(input_path: str, target_ext: str, output_dir: str, log=print) -> str:
    soffice = dependencies.find_soffice()
    if not soffice:
        raise ConversionError(
            "LibreOffice (soffice) tidak ditemukan. Install LibreOffice dan pastikan "
            "'soffice.exe' ada di PATH atau di 'C:\\Program Files\\LibreOffice\\program\\'."
        )
    os.makedirs(output_dir, exist_ok=True)
    cmd = [soffice, "--headless", "--norestore", "--convert-to", target_ext,
           "--outdir", output_dir, input_path]
    log(f"Menjalankan: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if result.returncode != 0:
        raise ConversionError(f"LibreOffice gagal: {result.stderr or result.stdout}")
    stem = Path(input_path).stem
    expected = os.path.join(output_dir, f"{stem}.{target_ext}")
    if not os.path.isfile(expected):
        raise ConversionError("Konversi LibreOffice selesai tapi file output tidak ditemukan.")
    return expected


def convert_txt_md_html(input_path: str, target_ext: str, output_path: str, log=print):
    """Konversi ringan antar txt/md/html murni Python (tanpa LibreOffice).
    HANYA dipakai kalau sumbernya memang berbasis teks polos (txt/md/html/rtf-text) -
    JANGAN dipakai untuk docx/doc/odt/pdf (format biner) karena hasilnya akan rusak."""
    PLAIN_TEXT_SOURCES = ("txt", "md", "html")
    src_ext = Path(input_path).suffix.lower().lstrip(".")
    if src_ext not in PLAIN_TEXT_SOURCES:
        return None  # bukan sumber teks polos -> serahkan ke LibreOffice/Word/fallback lain

    text = Path(input_path).read_text(encoding="utf-8", errors="ignore")

    if src_ext == "md" and target_ext == "html":
        try:
            import markdown
            html = markdown.markdown(text)
        except ImportError:
            html = f"<pre>{text}</pre>"
        Path(output_path).write_text(html, encoding="utf-8")
        return output_path

    if target_ext == "txt":
        # Strip tag html sederhana / biarkan md apa adanya sebagai txt
        Path(output_path).write_text(text, encoding="utf-8")
        return output_path

    return None  # fallback ke LibreOffice/Word untuk target lain (pdf, docx, rtf, dst)


# ---------------------------------------------------------------------------
# GAMBAR (Pillow)
# ---------------------------------------------------------------------------
def convert_image(input_path: str, target_ext: str, output_path: str, log=print):
    if Image is None:
        raise ConversionError("Pustaka Pillow belum terpasang. Jalankan: pip install pillow")

    src_ext = Path(input_path).suffix.lower().lstrip(".")
    if src_ext == "svg":
        try:
            import cairosvg
        except ImportError:
            raise ConversionError(
                "Konversi SVG butuh pustaka 'cairosvg'. Jalankan: pip install cairosvg"
            )
        if target_ext == "pdf":
            cairosvg.svg2pdf(url=input_path, write_to=output_path)
        else:
            png_tmp = output_path if target_ext == "png" else output_path + ".tmp.png"
            cairosvg.svg2png(url=input_path, write_to=png_tmp)
            if target_ext != "png":
                img = Image.open(png_tmp).convert("RGB")
                img.save(output_path)
                os.remove(png_tmp)
        return output_path

    img = Image.open(input_path)
    fmt_map = {"jpg": "JPEG", "jpeg": "JPEG", "png": "PNG", "webp": "WEBP",
               "bmp": "BMP", "tiff": "TIFF", "pdf": "PDF"}
    save_fmt = fmt_map.get(target_ext, target_ext.upper())

    if save_fmt in ("JPEG", "PDF", "BMP") and img.mode in ("RGBA", "P"):
        img = img.convert("RGB")

    img.save(output_path, save_fmt)
    return output_path


# ---------------------------------------------------------------------------
# AUDIO & VIDEO (FFmpeg)
# ---------------------------------------------------------------------------
def convert_media(input_path: str, target_ext: str, output_path: str, log=print):
    ffmpeg = dependencies.find_ffmpeg()
    if not ffmpeg:
        raise ConversionError(
            "FFmpeg tidak ditemukan. Download dari ffmpeg.org, lalu tambahkan folder "
            "'bin' FFmpeg ke PATH Windows."
        )
    cmd = [ffmpeg, "-y", "-i", input_path]
    # Codec sensible defaults
    codec_map = {
        "mp3": ["-c:a", "libmp3lame"],
        "aac": ["-c:a", "aac"],
        "opus": ["-c:a", "libopus"],
        "flac": ["-c:a", "flac"],
        "ogg": ["-c:a", "libvorbis"],
        "wav": ["-c:a", "pcm_s16le"],
        "m4a": ["-c:a", "aac"],
        "mp4": ["-c:v", "libx264", "-c:a", "aac"],
        "webm": ["-c:v", "libvpx-vp9", "-c:a", "libopus"],
        "mkv": ["-c:v", "libx264", "-c:a", "aac"],
        "mov": ["-c:v", "libx264", "-c:a", "aac"],
        "avi": ["-c:v", "mpeg4", "-c:a", "mp3"],
    }
    cmd += codec_map.get(target_ext, [])
    cmd.append(output_path)
    log(f"Menjalankan: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=1800)
    if result.returncode != 0:
        raise ConversionError(f"FFmpeg gagal: {result.stderr[-800:]}")
    return output_path


# ---------------------------------------------------------------------------
# ARSIP & KOMPRESI
# ---------------------------------------------------------------------------
def _extract_archive(input_path: str, extract_dir: str, log=print):
    src_ext = _archive_ext(input_path)
    os.makedirs(extract_dir, exist_ok=True)
    if src_ext == "zip":
        with zipfile.ZipFile(input_path) as z:
            z.extractall(extract_dir)
    elif src_ext in ("tar", "tar.gz", "tgz", "tar.bz2", "tar.xz"):
        mode = {"tar": "r:", "tar.gz": "r:gz", "tgz": "r:gz",
                "tar.bz2": "r:bz2", "tar.xz": "r:xz"}[src_ext]
        with tarfile.open(input_path, mode) as t:
            t.extractall(extract_dir)
    elif src_ext == "7z":
        if py7zr is None:
            raise ConversionError("Butuh pustaka 'py7zr'. Jalankan: pip install py7zr")
        with py7zr.SevenZipFile(input_path, mode="r") as z:
            z.extractall(extract_dir)
    elif src_ext == "rar":
        try:
            import rarfile
        except ImportError:
            raise ConversionError(
                "Butuh pustaka 'rarfile' + tool 'unrar'/'unar' terpasang di sistem "
                "untuk membuka file .rar. Jalankan: pip install rarfile"
            )
        with rarfile.RarFile(input_path) as z:
            z.extractall(extract_dir)
    else:
        raise ConversionError(f"Format arsip sumber '{src_ext}' tidak dikenali.")
    return extract_dir


def _archive_ext(path: str) -> str:
    name = Path(path).name.lower()
    for multi in ("tar.gz", "tar.bz2", "tar.xz"):
        if name.endswith("." + multi):
            return multi
    return Path(path).suffix.lower().lstrip(".")


def convert_archive(input_path: str, target_ext: str, output_path: str, log=print):
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        log("Mengekstrak arsip sumber...")
        extracted = _extract_archive(input_path, tmp, log=log)
        entries = list(Path(extracted).iterdir())
        # Jika hasil ekstrak hanya 1 folder, kompres isi folder itu langsung
        base = extracted

        log(f"Membuat arsip target ({target_ext})...")
        if target_ext == "zip":
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
                for root, _, files in os.walk(base):
                    for f in files:
                        full = os.path.join(root, f)
                        rel = os.path.relpath(full, base)
                        z.write(full, rel)
        elif target_ext in ("tar", "tar.gz", "tar.xz", "tar.bz2"):
            mode = {"tar": "w:", "tar.gz": "w:gz", "tar.xz": "w:xz", "tar.bz2": "w:bz2"}[target_ext]
            with tarfile.open(output_path, mode) as t:
                t.add(base, arcname=".")
        elif target_ext == "7z":
            if py7zr is None:
                raise ConversionError("Butuh pustaka 'py7zr'. Jalankan: pip install py7zr")
            with py7zr.SevenZipFile(output_path, "w") as z:
                z.writeall(base, ".")
        else:
            raise ConversionError(f"Format arsip tujuan '{target_ext}' tidak didukung.")
    return output_path


# ---------------------------------------------------------------------------
# SPREADSHEET & DATA TERSTRUKTUR
# ---------------------------------------------------------------------------
STRUCTURED_SOURCE_EXTS = ("json", "xml", "yaml", "yml")
STRUCTURED_TARGET_EXTS = ("json", "xml", "yaml", "yml", "toml")


def _read_tabular(input_path: str):
    ext = Path(input_path).suffix.lower().lstrip(".")

    if ext in STRUCTURED_SOURCE_EXTS:
        # Muat sebagai data terstruktur dulu, lalu coba jadikan tabel.
        data = _load_structured(input_path, ext)
        if isinstance(data, list):
            return pd.DataFrame(data)
        if isinstance(data, dict):
            # Kalau dict berisi satu list utama (pola umum export data), pakai itu.
            list_values = [v for v in data.values() if isinstance(v, list)]
            if len(data) == 1 and list_values:
                return pd.DataFrame(list_values[0])
            return pd.DataFrame([data])
        raise ConversionError("Struktur data tidak dapat direpresentasikan sebagai tabel.")

    if pd is None:
        raise ConversionError("Butuh pustaka 'pandas'. Jalankan: pip install pandas openpyxl")
    if ext in ("xlsx", "xls", "ods"):
        return pd.read_excel(input_path)
    if ext == "csv":
        return pd.read_csv(input_path)
    if ext == "tsv":
        return pd.read_csv(input_path, sep="\t")
    raise ConversionError(f"Format tabular '{ext}' belum didukung pembacaannya.")


def convert_spreadsheet(input_path: str, target_ext: str, output_path: str, log=print):
    src_ext = Path(input_path).suffix.lower().lstrip(".")

    # Sumber & tujuan sama-sama format data terstruktur murni (bukan tabel)
    # -> pertahankan struktur nested apa adanya, jangan dipaksa jadi tabel.
    if src_ext in STRUCTURED_SOURCE_EXTS and target_ext in STRUCTURED_TARGET_EXTS:
        return convert_structured_data(input_path, target_ext, output_path, log=log)

    # xlsx/xls/ods/pdf yang butuh render -> LibreOffice
    if target_ext in ("ods", "pdf"):
        return convert_via_libreoffice(input_path, target_ext, str(Path(output_path).parent), log=log)

    if pd is None:
        raise ConversionError("Butuh pustaka 'pandas'. Jalankan: pip install pandas openpyxl")

    df = _read_tabular(input_path)
    if target_ext == "csv":
        df.to_csv(output_path, index=False)
    elif target_ext == "xlsx":
        df.to_excel(output_path, index=False)
    elif target_ext == "txt":
        df.to_csv(output_path, index=False, sep="\t")
    elif target_ext == "json":
        df.to_json(output_path, orient="records", indent=2, force_ascii=False)
    elif target_ext == "html":
        df.to_html(output_path, index=False)
    elif target_ext == "md":
        Path(output_path).write_text(df.to_markdown(index=False), encoding="utf-8")
    else:
        raise ConversionError(f"Target spreadsheet '{target_ext}' tidak didukung.")
    return output_path


def convert_structured_data(input_path: str, target_ext: str, output_path: str, log=print):
    """json / yaml / xml / toml / ini / conf <-> saling konversi."""
    src_ext = Path(input_path).suffix.lower().lstrip(".")
    data = _load_structured(input_path, src_ext)
    _dump_structured(data, output_path, target_ext)
    return output_path


def _load_structured(path: str, ext: str):
    text = Path(path).read_text(encoding="utf-8")
    if ext == "json":
        return json.loads(text)
    if ext in ("yaml", "yml"):
        if yaml is None:
            raise ConversionError("Butuh pustaka 'pyyaml'. Jalankan: pip install pyyaml")
        return yaml.safe_load(text)
    if ext == "toml":
        if toml is None:
            raise ConversionError("Butuh pustaka 'toml'. Jalankan: pip install toml")
        return toml.loads(text)
    if ext in ("ini", "conf"):
        import configparser
        cp = configparser.ConfigParser()
        cp.read(path, encoding="utf-8")
        return {s: dict(cp.items(s)) for s in cp.sections()}
    if ext == "xml":
        import xml.etree.ElementTree as ET
        root = ET.fromstring(text)

        def node_to_dict(node):
            d = {}
            if node.attrib:
                d["@attributes"] = node.attrib
            children = list(node)
            if children:
                for c in children:
                    d.setdefault(c.tag, []).append(node_to_dict(c))
                for k, v in list(d.items()):
                    if isinstance(v, list) and len(v) == 1 and k != "@attributes":
                        d[k] = v[0]
            else:
                text_val = (node.text or "").strip()
                if text_val:
                    d["#text"] = text_val
            return d
        return {root.tag: node_to_dict(root)}
    raise ConversionError(f"Format sumber '{ext}' tidak dikenali untuk data terstruktur.")


def _dump_structured(data, output_path: str, ext: str):
    if ext == "json":
        Path(output_path).write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    elif ext in ("yaml", "yml"):
        if yaml is None:
            raise ConversionError("Butuh pustaka 'pyyaml'. Jalankan: pip install pyyaml")
        Path(output_path).write_text(yaml.dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")
    elif ext == "toml":
        if toml is None:
            raise ConversionError("Butuh pustaka 'toml'. Jalankan: pip install toml")
        if not isinstance(data, dict):
            raise ConversionError("TOML memerlukan struktur data berupa objek/dict di level atas.")
        Path(output_path).write_text(toml.dumps(data), encoding="utf-8")
    elif ext == "txt":
        Path(output_path).write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    elif ext == "csv":
        if not isinstance(data, list):
            raise ConversionError("CSV memerlukan struktur data berupa list of records.")
        if data and isinstance(data[0], dict):
            keys = list(data[0].keys())
            with open(output_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=keys)
                writer.writeheader()
                writer.writerows(data)
        else:
            with open(output_path, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerows(data)
    elif ext == "xml":
        import xml.etree.ElementTree as ET

        def build(parent, key, value):
            if isinstance(value, dict):
                el = ET.SubElement(parent, key)
                for k, v in value.items():
                    build(el, k, v)
            elif isinstance(value, list):
                for item in value:
                    build(parent, key, item)
            else:
                el = ET.SubElement(parent, key)
                el.text = str(value)

        root_key = "root"
        root = ET.Element(root_key)
        if isinstance(data, dict):
            for k, v in data.items():
                build(root, k, v)
        else:
            build(root, "item", data)
        ET.ElementTree(root).write(output_path, encoding="utf-8", xml_declaration=True)
    else:
        raise ConversionError(f"Format tujuan '{ext}' tidak didukung untuk data terstruktur.")


# ---------------------------------------------------------------------------
# DISPATCHER UTAMA
# ---------------------------------------------------------------------------
from . import format_map  # noqa: E402


def convert_file(input_path: str, target_ext: str, output_dir: str, log=print) -> str:
    """Titik masuk tunggal: tentukan kategori source lalu panggil converter yang sesuai."""
    src_ext = Path(input_path).suffix.lower().lstrip(".")
    # Tangani arsip multi-part ext (tar.gz dst)
    if src_ext == "gz" and Path(input_path).stem.endswith(".tar"):
        src_ext = "tar.gz"

    category = format_map.find_category(src_ext)
    if category is None:
        raise ConversionError(f"Format sumber '.{src_ext}' tidak dikenali oleh aplikasi.")

    os.makedirs(output_dir, exist_ok=True)
    stem = Path(input_path).stem
    if src_ext == "tar.gz":
        stem = Path(stem).stem  # buang .tar juga
    output_path = os.path.join(output_dir, f"{stem}.{target_ext}")

    log(f"[{category}] {os.path.basename(input_path)} -> .{target_ext}")

    if category == "Dokumen & Teks":
        light = convert_txt_md_html(input_path, target_ext, output_path, log=log)
        if light:
            return light

        soffice_available = dependencies.find_soffice() is not None
        word_available = dependencies.find_word() is not None

        if soffice_available:
            return convert_via_libreoffice(input_path, target_ext, output_dir, log=log)
        if word_available:
            log("LibreOffice tidak ditemukan, memakai Microsoft Word...")
            return convert_via_word(input_path, target_ext, output_dir, log=log)

        log("LibreOffice & Word tidak ditemukan, memakai mode ringan (teks polos)...")
        return convert_docx_lightweight(input_path, target_ext, output_path, log=log)

    if category == "Presentasi":
        soffice_available = dependencies.find_soffice() is not None
        if soffice_available:
            return convert_via_libreoffice(input_path, target_ext, output_dir, log=log)
        raise ConversionError(
            "Konversi presentasi (pptx/ppt/odp) butuh LibreOffice terpasang. "
            "Mode ringan Python belum mendukung format presentasi."
        )

    if category == "Spreadsheet & Data":
        return convert_spreadsheet(input_path, target_ext, output_path, log=log)

    if category == "Gambar":
        return convert_image(input_path, target_ext, output_path, log=log)

    if category == "Audio":
        return convert_media(input_path, target_ext, output_path, log=log)

    if category == "Video":
        return convert_media(input_path, target_ext, output_path, log=log)

    if category == "Arsip & Kompresi":
        return convert_archive(input_path, target_ext, output_path, log=log)

    if category == "Data Terstruktur (Kode & Konfigurasi)":
        return convert_structured_data(input_path, target_ext, output_path, log=log)

    raise ConversionError(f"Kategori '{category}' belum diimplementasikan.")
