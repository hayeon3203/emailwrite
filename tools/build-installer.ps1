# publish -> ISCC. publish는 앱이 떠 있으면 MSB3027로 죽는다(EmailWrite.Core.dll이 잠긴다).
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

Get-Process EmailWrite.App -ErrorAction SilentlyContinue | Stop-Process -Force

dotnet publish "$root\src\EmailWrite.App" -c Release -r win-x64 --self-contained -o "$root\publish\EmailWrite"
if ($LASTEXITCODE -ne 0) { throw "publish 실패" }

$iscc = "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $iscc)) { throw "Inno Setup이 없다: $iscc (winget install JRSoftware.InnoSetup)" }

& $iscc "$root\installer\EmailWrite.iss"
if ($LASTEXITCODE -ne 0) { throw "ISCC 실패" }

Get-ChildItem "$root\installer\Output\*.exe" | ForEach-Object {
    "{0}  {1:N0} MB" -f $_.FullName, ($_.Length / 1MB)
}
