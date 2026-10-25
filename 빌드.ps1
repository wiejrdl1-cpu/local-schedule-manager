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
    & $BuildPython -c "import PyInstaller, cryptography, holidays, openpyxl, pandas, xlrd, windnd"
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
        --runtime-tmpdir 'C:\Users\Public\Documents\TaskItEasyTemp' `
        --name 'Task It Easy' `
        --icon 'assets\app_icon.ico' `
        --add-data 'assets\app_icon.ico;assets' `
        --add-data 'assets\fonts\Fredoka-Bold.ttf;assets\fonts' `
        --add-data 'assets\fonts\OFL-Fredoka.txt;assets\fonts' `
        --additional-hooks-dir 'build_hooks' `
        --collect-data 'holidays' `
        --paths '.vendor' `
        --hidden-import 'xlrd' `
        --hidden-import 'windnd' `
        app.py
    if ($LASTEXITCODE -ne 0) {
        throw "EXE build failed with exit code $LASTEXITCODE"
    }
} finally {
    Pop-Location
}

$ExePath = Join-Path $ProjectDir 'dist\Task It Easy.exe'
if (-not (Test-Path -LiteralPath $ExePath)) {
    throw 'The built EXE file was not found.'
}

$ArchiveViewer = Join-Path $ProjectDir '.venv-build\Scripts\pyi-archive_viewer.exe'
$ArchiveListing = & $ArchiveViewer -r -b $ExePath
$TkinterIncluded = (
    $LASTEXITCODE -eq 0 -and
    $ArchiveListing -match '(?m)^ _tkinter\.pyd$' -and
    $ArchiveListing -match '(?m)^ tkinter$' -and
    $ArchiveListing -match '(?m)^ _tcl_data\\init\.tcl$' -and
    $ArchiveListing -match '(?m)^ _tk_data\\tk\.tcl$'
)
if (-not $TkinterIncluded) {
    throw 'EXE verification failed: tkinter was not bundled.'
}

Write-Host "EXE build completed: $ExePath"
