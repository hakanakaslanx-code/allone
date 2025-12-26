; Inno Setup script for AllOne stable launcher + release layout.

#define AppName "AllOne"
#define AppVersion "0.0.0"
#define AppPublisher "AllOne"
#define AppExeName "AllOneLauncher.exe"
#define ReleaseExeName "AllOne Tools.exe"

[Setup]
AppId={{9B0F7F71-9850-4F22-9F97-7A2B0E9E18A1}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\AllOne
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
Compression=lzma
SolidCompression=yes
OutputBaseFilename=AllOneSetup
UninstallDisplayIcon={app}\{#AppExeName}

[Files]
Source: "dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\releases\{#AppVersion}\*"; DestDir: "{app}\releases\{#AppVersion}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent
