; 설치 파일 스크립트. 컴파일은 tools\build-installer.ps1 (publish 먼저 돌린다).
;
; 관리자 권한을 쓰지 않는다(PrivilegesRequired=lowest, {autopf} = %LocalAppData%\Programs).
; 자동 시작이 이미 HKCU Run 값 하나로 사용자별이고, 동료 PC에 관리자 권한이 있다는 보장이
; 없다. 모델이 들어가는 {commonappdata}(C:\ProgramData)는 Users가 승격 없이 쓸 수 있다.

#define AppExe "EmailWrite.App.exe"
#define SourceDir "..\publish\EmailWrite"
#define ModelFile "hyperclovax-seed-text-instruct-1.5b-q4_k_m.gguf"
; 컴파일하는 PC의 C:\ProgramData\EmailWrite\Models 에서 집어넣는다 — 앱이 받아둔 그 파일이다.
#define ModelSource GetEnv("ProgramData") + "\EmailWrite\Models\" + ModelFile
; exe에서 읽는다 — .iss에 손으로 적으면 csproj의 <Version>과 갈라진다.
#define AppVersion GetVersionNumbersString(SourceDir + "\" + AppExe)

[Setup]
AppId={{7E1C4F0A-3B2D-4E58-9C7A-5D6B8E2F1A34}
AppName=EmailWrite
AppVersion={#AppVersion}
AppPublisher=다올티에스
DefaultDirName={autopf}\EmailWrite
DefaultGroupName=EmailWrite
PrivilegesRequired=lowest
; 실행 중이면 파일이 잠긴다. App.xaml.cs의 인스턴스 뮤텍스와 같은 이름 — 설치/제거 전에
; 닫아달라는 창이 뜬다.
AppMutex=EmailWrite.App
; 재배포에 가중치가 들어가므로 HyperCLOVA X SEED 3.1(ii)의 "약관 사본 전달"을 설치 화면에서도 만족시킨다.
LicenseFile=..\src\EmailWrite.App\LICENSE-HyperCLOVAX.txt
SetupIconFile=..\src\EmailWrite.App\app.ico
UninstallDisplayIcon={app}\{#AppExe}
; 기본값은 "EmailWrite 버전 1.0.0.0"이다. 설치된 앱 목록에 이름만 두고, 버전은 그 항목의
; 정보(DisplayVersion)로만 보이게 한다.
UninstallDisplayName=EmailWrite
OutputDir=Output
OutputBaseFilename=EmailWrite-{#AppVersion}-setup
; gguf는 이미 4비트라 압축이 거의 안 먹는다. solid로 묶으면 컴파일만 몇 배 느려진다.
Compression=lzma2/fast
SolidCompression=no
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕 화면에 바로 가기 만들기"; Flags: unchecked

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; 첫 실행 다운로드(1.06GB)를 안 태운다. LocalDraftGenerator가 File.Exists면 건너뛰므로 코드는 그대로다.
; onlyifdoesntexist — 이미 받아둔 PC의 재설치에서 1GB를 다시 쓰지 않는다.
Source: "{#ModelSource}"; DestDir: "{commonappdata}\EmailWrite\Models"; Flags: onlyifdoesntexist nocompression

[Icons]
Name: "{group}\EmailWrite"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\EmailWrite"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Registry]
; 자동 시작은 앱이 켜고 끈다(Autostart.cs). 제거할 때 값만 지운다 — 안 지우면 없는 exe를
; 가리키는 Run 항목이 남는다. dontcreatekey: 설치가 자동 시작을 켜지는 않는다.
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueName: "EmailWrite"; Flags: dontcreatekey uninsdeletevalue

[Run]
Filename: "{app}\{#AppExe}"; Description: "EmailWrite 실행"; Flags: nowait postinstall skipifsilent
