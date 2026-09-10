#define AppName "中文文本词频统计分析工具"
#define AppVersion "4.2.0"
#define AppPublisher "陈浩杰｜澳门城市大学金融学院"
#define AppExeName "词频统计分析工具.exe"

[Setup]
AppId={{B99E2E50-4E92-4F15-9D9B-6F5F1B0D6B20}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
OutputDir=..\dist\release
OutputBaseFilename=词频统计工具-Windows-x64-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\{#AppExeName}
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE

[Files]
Source: "..\dist\词频统计分析工具\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"

[Run]
Filename: "{app}\{#AppExeName}"; Description: "启动 {#AppName}"; Flags: postinstall nowait skipifsilent
