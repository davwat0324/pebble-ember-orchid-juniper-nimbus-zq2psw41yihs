@echo off
rem Double-click to install Remind for the current Windows user (no admin rights needed).
rem PowerShell runs hidden; Install.ps1 reports in a message box. It is run as text (not as a script
rem file) so an execution policy set by Group Policy cannot stop it, and any error still shows a box.
set "REMIND_DIR=%~dp0"
if exist "%~dp0app\Remind.exe" goto install
powershell -NoProfile -WindowStyle Hidden -Command "Add-Type -AssemblyName PresentationFramework; [void][System.Windows.MessageBox]::Show('Unzip Remind-win64.zip first: right-click it and choose Extract All. Then double-click Install.cmd in the extracted folder.', 'Remind')"
exit /b 1
:install
start "" powershell -NoProfile -WindowStyle Hidden -Command "try { & ([scriptblock]::Create((Get-Content -Raw -LiteralPath ($env:REMIND_DIR + 'Install.ps1')))) -Source ($env:REMIND_DIR + 'app') } catch { Add-Type -AssemblyName PresentationFramework; [void][System.Windows.MessageBox]::Show('Remind could not be installed: ' + $_.Exception.Message, 'Remind') }"
