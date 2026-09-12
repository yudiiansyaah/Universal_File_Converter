@echo off
REM ============================================================
REM  Build Universal File Converter menjadi file .exe (Windows)
REM  Jalankan file ini di Command Prompt / PowerShell pada folder
REM  proyek ini, SETELAH menjalankan:
REM      python -m venv venv
REM      venv\Scripts\activate
REM      pip install -r requirements.txt
REM ============================================================

echo Membersihkan build lama...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q UniversalFileConverter.spec 2>nul

echo Membuat executable dengan PyInstaller...
pyinstaller --noconfirm --windowed --onefile ^
    --name "UniversalFileConverter" ^
    --icon "assets\icon.ico" ^
    --collect-all PySide6 ^
    main.py

if not exist "dist\UniversalFileConverter.exe" (
    echo.
    echo [GAGAL] File exe tidak ditemukan. Cek pesan error PyInstaller di atas.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo Build .exe selesai: dist\UniversalFileConverter.exe
echo ============================================================
echo.

REM ----- Coba build installer otomatis kalau Inno Setup terpasang -----
set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"

if exist %ISCC% (
    echo Inno Setup terdeteksi, membuat installer...
    %ISCC% installer.iss
    echo.
    echo ============================================================
    echo Installer selesai: installer_output\UniversalFileConverterSetup.exe
    echo File inilah yang dibagikan ke user awam - tinggal Next-Next-Install.
    echo ============================================================
) else (
    echo.
    echo [INFO] Inno Setup belum terdeteksi, installer belum dibuat.
    echo Download gratis di https://jrsoftware.org/isinfo.php lalu jalankan:
    echo   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
    echo Atau cukup bagikan dist\UniversalFileConverter.exe langsung ^(tanpa installer^).
)
pause
