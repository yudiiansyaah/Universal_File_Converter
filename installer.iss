; ============================================================
;  Installer script untuk Universal File Converter
;  Compile dengan Inno Setup (https://jrsoftware.org/isinfo.php)
;  Hasil: UniversalFileConverterSetup.exe
;
;  CARA PAKAI:
;  1. Download & install Inno Setup (gratis) dari jrsoftware.org
;  2. Pastikan dist\UniversalFileConverter.exe sudah ada
;     (jalankan build.bat dulu)
;  3. Buka file ini dengan Inno Setup Compiler, klik Compile (F9)
;     ATAU jalankan dari command line:
;         "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
; ============================================================

#define MyAppName "Universal File Converter"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Yudiansyah"
#define MyAppExeName "UniversalFileConverter.exe"

[Setup]
AppId={{B4E1A6F2-7C3D-4A9B-9F1E-8D2C6A5B3E10}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=UniversalFileConverterSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=assets\icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Buat shortcut di Desktop"; GroupDescription: "Shortcut tambahan:"; Flags: unchecked

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Jalankan {#MyAppName} sekarang"; Flags: nowait postinstall skipifsilent
