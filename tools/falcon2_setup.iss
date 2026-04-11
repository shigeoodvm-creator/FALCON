; FALCON2 Inno Setup スクリプト
; Inno Setup 6.x 以上が必要
; ダウンロード: https://jrsoftware.org/isinfo.php

#define AppName      "FALCON2"
#define AppVersion   "1.0.0"
#define AppPublisher "FALCON開発チーム"
#define AppExeName   "FALCON2.exe"
#define SourceDir    "..\dist\FALCON2"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisherURL=
AppSupportURL=
AppUpdatesURL=
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
LicenseFile=
OutputDir=..\dist\installer
OutputBaseFilename=FALCON2_Setup_v{#AppVersion}
SetupIconFile=..\app\resources\falcon2.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; 64bit Windows のみ対応
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; 管理者権限不要（ユーザーフォルダにもインストール可）
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; アンインストール設定
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName} v{#AppVersion}

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにショートカットを作成"; GroupDescription: "追加タスク:"

[Files]
; EXE 本体
Source: "{#SourceDir}\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; ライブラリ群（_internal フォルダごと）
Source: "{#SourceDir}\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; スタートメニュー
Name: "{group}\{#AppName}";          Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"
Name: "{group}\{#AppName} のアンインストール"; Filename: "{uninstallexe}"
; デスクトップ（タスク選択時のみ）
Name: "{autodesktop}\{#AppName}";    Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; インストール完了後に起動するか確認
Filename: "{app}\{#AppExeName}"; Description: "FALCON2 を今すぐ起動する"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; アンインストール時に config/ フォルダも削除（任意）
; ユーザーデータ (C:\FARMS) は削除しない
Type: filesandordirs; Name: "{app}\config"
