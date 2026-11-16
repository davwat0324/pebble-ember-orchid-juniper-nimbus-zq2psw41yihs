<#
  Builds dist\Remind-win64.zip: the PyInstaller app (Remind-win64\app\Remind.exe + _internal) plus
  Install.cmd / Install.ps1. Used by the GitHub Actions release workflow and for local test builds.

  Needs Python 3.14 with: pip install -r requirements.txt
#>
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$src = Join-Path $root 'src'
$build = Join-Path $root 'build'
$dist = Join-Path $root 'dist'
Remove-Item $build, $dist -Recurse -Force -ErrorAction SilentlyContinue

$ico = Join-Path $src 'remind.ico'
# openpyxl and PySide6 pull in optional heavy packages Remind never uses; leave them out.
$exclude = 'numpy', 'PIL', 'pandas', 'tkinter', 'matplotlib', 'PySide6.QtQml', 'PySide6.QtQuick',
           'PySide6.QtQuickWidgets', 'PySide6.QtPdf', 'PySide6.QtNetwork', 'PySide6.QtOpenGL',
           'PySide6.QtOpenGLWidgets', 'PySide6.QtSvg' | ForEach-Object { '--exclude-module'; $_ }
python -m PyInstaller --noconfirm --clean --windowed --name Remind `
    --icon $ico --add-data "$ico;." @exclude `
    --distpath (Join-Path $build 'pyi-dist') --workpath (Join-Path $build 'pyi-work') --specpath $build `
    (Join-Path $src 'main.py')
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed ($LASTEXITCODE)" }

# Qt's software OpenGL fallback (20 MB) is only for 3D/Quick content; plain widgets don't need it.
Remove-Item (Join-Path $build 'pyi-dist\Remind\_internal\PySide6\opengl32sw.dll') -ErrorAction SilentlyContinue

$stage = Join-Path $build 'Remind-win64'
New-Item -ItemType Directory -Force $stage | Out-Null
Move-Item (Join-Path $build 'pyi-dist\Remind') (Join-Path $stage 'app')
Copy-Item (Join-Path $root 'installer\Install.cmd'), (Join-Path $root 'installer\Install.ps1') $stage

New-Item -ItemType Directory -Force $dist | Out-Null
$zip = Join-Path $dist 'Remind-win64.zip'
Compress-Archive -Path $stage -DestinationPath $zip -CompressionLevel Optimal
Write-Host "Built $zip ($([math]::Round((Get-Item $zip).Length / 1MB, 1)) MB)"
