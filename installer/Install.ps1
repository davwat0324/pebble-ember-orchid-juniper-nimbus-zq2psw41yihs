<#
  Installs Remind for the current Windows user: copies the app to %LOCALAPPDATA%\Programs\Remind, adds
  Desktop + Start-menu shortcuts and starts it. No admin rights needed. Install.cmd runs it hidden
  (passing -Source <unzipped folder>\app) and it reports in a message box. Settings, the saved Spectrum login and history (%APPDATA%\Remind) are not touched.
#>
param([Parameter(Mandatory)][string]$Source)
Add-Type -AssemblyName PresentationFramework
$install = Join-Path $env:LOCALAPPDATA 'Programs\Remind'
$exe = Join-Path $install 'Remind.exe'
try {
    $ErrorActionPreference = 'Stop'
    Get-Process Remind -ErrorAction SilentlyContinue | Stop-Process -Force -PassThru | Wait-Process -Timeout 10 -ErrorAction SilentlyContinue
    robocopy $Source $install /MIR /R:3 /W:2 /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "Copying to $install failed." }
    Get-ChildItem $install -Recurse -File | Unblock-File      # clears the "downloaded from the internet" mark

    $shell = New-Object -ComObject WScript.Shell
    foreach ($dir in [Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs')) {
        $s = $shell.CreateShortcut((Join-Path $dir 'Remind.lnk'))
        $s.TargetPath = $exe
        $s.WorkingDirectory = $install
        $s.Description = 'Remind - email reminders to employees'
        $s.Save()
    }

    $notes = @()
    if (-not (Get-OdbcDriver -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^ODBC Driver (17|18) for SQL Server$' })) {
        $notes += 'Microsoft ODBC Driver 18 for SQL Server is not installed. Remind needs it to read emails from Spectrum - ask IT to install it (free from Microsoft).'
    }
    if (-not (Test-Path 'Registry::HKEY_CLASSES_ROOT\Outlook.Application')) {
        $notes += 'Classic Outlook was not found. Remind sends through classic (desktop) Outlook.'
    }
    $msg = "Remind is installed. It is on your Desktop and in the Start menu."
    if ($notes) { $msg += "`n`nBefore you use it:`n- " + ($notes -join "`n- ") }
    [void][System.Windows.MessageBox]::Show($msg, 'Remind', 'OK', $(if ($notes) { 'Warning' } else { 'Information' }))
    Start-Process $exe
} catch {
    [void][System.Windows.MessageBox]::Show("Remind could not be installed.`n`n$($_.Exception.Message)", 'Remind', 'OK', 'Error')
    exit 1
}
