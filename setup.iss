#define AppName "Outlook Draft Mailer"
#define AppVersion "1.0.0"
#define AppPublisher "Outlook Draft Mailer"

[Setup]
AppId={{A4F5BE86-0D1C-4DC4-A8B9-6642B0C7FA91}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\Outlook Draft Mailer
DefaultGroupName={#AppName}
UninstallDisplayIcon={app}\Outlook Draft Mailer.exe
OutputDir=dist\installer
OutputBaseFilename=OutlookDraftMailerSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "dist\Outlook Draft Mailer\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\Outlook Draft Mailer.exe"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\Outlook Draft Mailer.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Outlook Draft Mailer.exe"; Description: "Launch {#AppName}"; Flags: postinstall nowait skipifsilent
