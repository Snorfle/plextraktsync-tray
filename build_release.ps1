param(
    [string]$Version = "0.1.0"
)

$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvDir = Join-Path $baseDir ".venv"
$pythonExe = Join-Path $venvDir "Scripts\python.exe"
$requirements = Join-Path $baseDir "requirements.txt"
$appScript = Join-Path $baseDir "plextraktsync_tray.py"
$releaseDir = Join-Path $baseDir "release"
$pyinstallerWorkDir = Join-Path $releaseDir "pyinstaller-build"
$pyinstallerDistDir = Join-Path $releaseDir "pyinstaller-dist"
$distAppDir = Join-Path $pyinstallerDistDir "PlexTraktSyncTray"
$stagingDir = Join-Path $releaseDir "PlexTraktSyncTray-$Version"
$zipPath = Join-Path $releaseDir "PlexTraktSyncTray-$Version.zip"

if (-not (Test-Path $venvDir)) {
    py -3 -m venv $venvDir
}

& $pythonExe -m pip install --upgrade pip
& $pythonExe -m pip install -r $requirements
& $pythonExe -m pip install pyinstaller

New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
& $pythonExe -m PyInstaller --noconfirm --windowed --name PlexTraktSyncTray --distpath $pyinstallerDistDir --workpath $pyinstallerWorkDir --specpath $releaseDir $appScript

if (Test-Path $stagingDir) {
    Remove-Item -LiteralPath $stagingDir -Recurse -Force
}
if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

New-Item -ItemType Directory -Path $stagingDir | Out-Null
Copy-Item -LiteralPath $distAppDir -Destination (Join-Path $stagingDir "PlexTraktSyncTray") -Recurse
Copy-Item -LiteralPath (Join-Path $baseDir "install_release.ps1") -Destination $stagingDir
Copy-Item -LiteralPath (Join-Path $baseDir "README.md") -Destination $stagingDir

Compress-Archive -Path (Join-Path $stagingDir "*") -DestinationPath $zipPath -Force

Write-Host "Built release zip: $zipPath"
