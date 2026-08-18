#define AppName "Book to LaTeX"
#define AppVersion "1.2.0"
#define AppExeName "book-reader.exe"

[Setup]
AppId={{D1CA797E-48B6-42D0-B1BA-E1FC91EF9940}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Book to LaTeX contributors
DefaultDirName={autopf}\Book to LaTeX
DefaultGroupName=Book to LaTeX
DisableProgramGroupPage=yes
OutputDir=..\..\dist
OutputBaseFilename=Book-to-LaTeX-Setup-{#AppVersion}-Windows-x64
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\{#AppExeName}

[Files]
Source: "..\..\dist\book-reader.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\USER_GUIDE.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Book to LaTeX"; Filename: "{app}\{#AppExeName}"
Name: "{userdesktop}\Book to LaTeX"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Open Book to LaTeX"; Flags: nowait postinstall skipifsilent
