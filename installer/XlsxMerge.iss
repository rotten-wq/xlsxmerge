; ============================================================
; XlsxMerge Inno Setup Script
; ============================================================
; 빌드: iscc.exe XlsxMerge.iss
; CI에서 /DVERSION=0.7.0 /DPUBLISH_DIR=..\publish\win-x64 로 오버라이드
; ============================================================

#ifndef VERSION
  #define VERSION "0.0.0"
#endif

#ifndef PUBLISH_DIR
  #define PUBLISH_DIR "..\publish\win-x64"
#endif

#define APP_NAME    "XlsxMerge"
#define APP_EXE     "XlsxMerge.exe"
#define PUBLISHER   "Nexon Korea"
#define URL         "https://github.com/rotten-wq/xlsxmerge"

[Setup]
AppId={{A7F2D3E1-8B4C-4F9A-B5D6-7E1C3A9F2B8D}
AppName={#APP_NAME}
AppVersion={#VERSION}
AppPublisher={#PUBLISHER}
AppPublisherURL={#URL}
AppSupportURL={#URL}/issues
DefaultDirName={autopf}\{#APP_NAME}
DefaultGroupName={#APP_NAME}
OutputDir=output
OutputBaseFilename=XlsxMerge-{#VERSION}-setup
SetupIconFile=..\src\XlsxMerge\XlsxMerge.ico
UninstallDisplayIcon={app}\{#APP_EXE}
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
LicenseFile=
; 라이선스가 필요하면 위에 경로 추가

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";   Description: "Create a &desktop shortcut";       GroupDescription: "Additional shortcuts:"
Name: "forkregister";  Description: "Register as Fork diff/merge tool"; GroupDescription: "Git Integration:"

[Files]
; 빌드 출력 전체 복사 (exe, dll, diff3.exe 등)
Source: "{#PUBLISH_DIR}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Fork 등록 스크립트
Source: "..\scripts\Register-ForkDiffTool.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#APP_NAME}";           Filename: "{app}\{#APP_EXE}"
Name: "{group}\Uninstall {#APP_NAME}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#APP_NAME}";     Filename: "{app}\{#APP_EXE}"; Tasks: desktopicon

[Run]
; 설치 후 Fork 등록 (사용자가 체크한 경우)
Filename: "powershell.exe"; \
  Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\Register-ForkDiffTool.ps1"" -ExePath ""{app}\{#APP_EXE}"" -Silent"; \
  StatusMsg: "Registering XlsxMerge in Fork..."; \
  Tasks: forkregister; \
  Flags: runhidden waituntilterminated

; 설치 완료 후 실행 옵션
Filename: "{app}\{#APP_EXE}"; \
  Description: "Launch {#APP_NAME}"; \
  Flags: nowait postinstall skipifsilent unchecked

[UninstallRun]
; 제거 시 Fork 등록 해제
Filename: "powershell.exe"; \
  Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\Register-ForkDiffTool.ps1"" -Uninstall -Silent"; \
  RunOnceId: "ForkUnregister"; \
  Flags: runhidden waituntilterminated

