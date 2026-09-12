"""
dependencies.py
Deteksi ketersediaan tool eksternal yang dibutuhkan (FFmpeg, LibreOffice).
"""
import shutil
import subprocess


def find_soffice() -> str | None:
    """Cari executable LibreOffice (soffice) di PATH atau lokasi umum Windows."""
    candidates = [
        "soffice",
        "soffice.exe",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        path = shutil.which(c)
        if path:
            return path
    import os
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def find_word() -> str | None:
    """Cari instalasi Microsoft Word (untuk automation via COM, Windows only)."""
    import os
    candidates = [
        r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    # Coba deteksi lewat registry (lebih akurat untuk versi Office apa pun)
    try:
        import winreg
        for hive, key in [
            (winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer"),
        ]:
            try:
                with winreg.OpenKey(hive, key) as k:
                    val, _ = winreg.QueryValueEx(k, "")
                    if val:
                        return "Word.Application"  # ProgID valid, COM bisa dipakai
            except FileNotFoundError:
                continue
    except ImportError:
        pass  # bukan Windows
    return None


def find_ffmpeg() -> str | None:
    return shutil.which("ffmpeg") or shutil.which("ffmpeg.exe")


def check_all() -> dict:
    """Kembalikan status semua dependensi eksternal."""
    return {
        "soffice": find_soffice(),
        "ffmpeg": find_ffmpeg(),
        "word": find_word(),
    }


def version_string(path: str) -> str:
    try:
        out = subprocess.run([path, "--version"], capture_output=True, text=True, timeout=10)
        return out.stdout.strip().splitlines()[0] if out.stdout else "OK"
    except Exception:
        return "OK"
