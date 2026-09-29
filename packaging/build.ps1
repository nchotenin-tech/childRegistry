param([string]$Python = '.\.venv\Scripts\python.exe', [string]$Iscc = '.\.build-tools\Inno\ISCC.exe')
$ErrorActionPreference = 'Stop'
Set-Location (Split-Path $PSScriptRoot -Parent)
$testDirectory = Join-Path $PWD ('.test-build-' + [guid]::NewGuid().ToString('N'))
& $Python -m pytest -q -p no:cacheprovider --basetemp $testDirectory
if ($LASTEXITCODE) { throw 'Tests failed' }
& $Python -m PyInstaller --noconfirm packaging/ChildRegistry.spec
if ($LASTEXITCODE) { throw 'PyInstaller failed' }
$env:QT_QPA_PLATFORM='offscreen'
$smoke = Start-Process -FilePath "$PWD\dist\ChildRegistry\ChildRegistry.exe" -ArgumentList '--smoke-test',"$PWD\tmp\frozen-check.json" -WindowStyle Hidden -PassThru
if (!$smoke.WaitForExit(60000)) { Stop-Process -Id $smoke.Id; throw 'Frozen smoke test timed out' }
if ($smoke.ExitCode -ne 0 -or !(Test-Path tmp/frozen-check.json)) { throw 'Frozen smoke test failed' }
& $Iscc packaging/installer.iss
if ($LASTEXITCODE) { throw 'Installer compilation failed' }
Get-FileHash installer-output/ChildRegistry-Setup-1.0.1.exe -Algorithm SHA256 | Format-List | Out-File installer-output/SHA256.txt
