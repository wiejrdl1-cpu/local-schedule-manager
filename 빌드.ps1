param(
    [switch]$InstallDependencies
)

$ErrorActionPreference = 'Stop'
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildPython = Join-Path $ProjectDir '.venv-build\Scripts\python.exe'
$BundledPython = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'

if (-not (Test-Path -LiteralPath $BuildPython)) {
    if (Test-Path -LiteralPath $BundledPython) {
        $BasePython = $BundledPython
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        $BasePython = (Get-Command python).Source
    } else {
        throw 'Python 3.11 or later is required.'
    }
    & $BasePython -m venv (Join-Path $ProjectDir '.venv-build')
}

if ($InstallDependencies) {
    & $BuildPython -m pip install --upgrade pip
    & $BuildPython -m pip install -r (Join-Path $ProjectDir 'requirements.txt') pyinstaller
}

try {
    & $BuildPython -c "import PyInstaller, cryptography, openpyxl, pandas, xlrd"
} catch {
    throw 'Build dependencies are missing. Run this script with -InstallDependencies first.'
}

Push-Location $ProjectDir
try {
    & $BuildPython -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --name 'LocalScheduleManager' `
        --paths '.vendor' `
        --hidden-import 'xlrd' `
        app.py
    if ($LASTEXITCODE -ne 0) {
        throw "EXE build failed with exit code $LASTEXITCODE"
    }
} finally {
    Pop-Location
}

$ExePath = Join-Path $ProjectDir 'dist\LocalScheduleManager.exe'
if (-not (Test-Path -LiteralPath $ExePath)) {
    throw 'The built EXE file was not found.'
}

Write-Host "EXE build completed: $ExePath"
