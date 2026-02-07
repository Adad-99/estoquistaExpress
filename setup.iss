; Script Inno Setup para Estoquista Express
; Compilar com Inno Setup Compiler (https://jrsoftware.org/isinfo.php)

#define MyAppName "Estoquista Express"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Guilherme Adad"
#define MyAppURL "https://github.com/Adad-99"
#define MyAppExeName "Estoquista Express.exe"

[Setup]
; Informações básicas
AppId={{E8F4A5B2-3C1D-4E9F-A7B6-9D2E5F8C1A3B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer
OutputBaseFilename=Estoquista_Express_Setup
SetupIconFile=resources\icone.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Executável principal
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Pasta resources da dist (se você copiou para lá)
Source: "dist\resources\*"; DestDir: "{app}\resources"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: DirExists('dist\resources')
; Pasta resources do projeto (fallback caso não esteja na dist)
Source: "resources\*"; DestDir: "{app}\resources"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: not DirExists('dist\resources')

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
