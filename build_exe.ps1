param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
$requirementsPath = Join-Path $projectRoot "requirements.txt"
$specPath = Join-Path $projectRoot "MacroRecorder.spec"
$distPath = Join-Path $projectRoot "dist"
$buildPath = Join-Path $projectRoot "build"

if (-not (Test-Path $venvPython)) {
    throw "Virtual environment Python was not found at '$venvPython'. Create the venv first."
}

Set-Location $projectRoot

if ($Clean) {
    Remove-Item -Recurse -Force $buildPath -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $distPath -ErrorAction SilentlyContinue
}

& $venvPython -m pip install -r $requirementsPath

if (Test-Path $specPath) {
    & $venvPython -m PyInstaller --noconfirm $specPath
}
else {
    & $venvPython -m PyInstaller --noconfirm --onefile --windowed --name "MacroRecorder" main.py
}

$exePath = Join-Path $distPath "MacroRecorder.exe"
if (Test-Path $exePath) {
    Write-Host "Build complete: $exePath"
}
else {
    Write-Warning "PyInstaller finished, but '$exePath' was not found. Check the build output above."
}
