; Inno Setup script for zerobounce
; Defines passed from build.bat: AppName, AppVersion, Publisher, SourceDir, OutputDir, OutputBaseFilename

#define MyAppURL "https://zerobounce.app"

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#Publisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
OutputDir={#OutputDir}
OutputBaseFilename={#OutputBaseFilename}
Compression=lzma
SolidCompression=yes
UninstallDisplayIcon={app}\zerobounce.exe
DisableProgramGroupPage=yes
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\zerobounce.exe"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\zerobounce.exe"

[Run]
Filename: "{app}\zerobounce.exe"; Description: "Launch {#AppName}"; Flags: postinstall nowait skipifsilent
