; 生产计划协同管理平台 安装脚本（Inno Setup 6）
; 编译：ISCC.exe installer.iss

#define MyAppName "生产计划协同管理平台"
#define MyAppVersion "1.4.0"
#define MyAppPublisher "内部生产系统"
#define MyAppExeName "生产计划协同管理平台.exe"
#define MyDistDir "..\dist\生产计划协同管理平台"

[Setup]
AppId={{9F3E7C2A-6B14-4D58-9A0C-2E5B7D1C8F40}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableDirPage=no
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
SetupIconFile=app.ico
OutputDir=..\dist
OutputBaseFilename={#MyAppName}_安装程序_v{#MyAppVersion}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; 默认装到当前用户目录（可写），保证 data 目录就落在安装目录下
PrivilegesRequired=lowest
UsePreviousAppDir=no
CloseApplications=yes
RestartApplications=no
AllowNoIcons=yes

[Languages]
Name: "chinesesimp"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务："; Flags: checkedonce

[Files]
Source: "{#MyDistDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent
