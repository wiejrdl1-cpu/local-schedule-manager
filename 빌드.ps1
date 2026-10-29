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

$SplashPath = Join-Path $ProjectDir 'assets\splash.png'
if (-not (Test-Path -LiteralPath $SplashPath)) {
    throw 'Splash image is missing. Run assets\create_splash.ps1 first.'
}

$VenvConfig = Join-Path $ProjectDir '.venv-build\pyvenv.cfg'
$PythonHomeLine = Get-Content -LiteralPath $VenvConfig -Encoding UTF8 | Where-Object { $_ -match '^home\s*=' } | Select-Object -First 1
if (-not $PythonHomeLine) {
    throw 'Build Python home could not be determined.'
}
$PythonHome = ($PythonHomeLine -replace '^home\s*=\s*', '').Trim()
$SourceTclRoot = Join-Path $PythonHome 'tcl'
$BuildTclRoot = 'C:\Users\Public\Documents\TaskItEasyBuildTcl'
if (-not (Test-Path -LiteralPath (Join-Path $SourceTclRoot 'tcl8.6')) -or
    -not (Test-Path -LiteralPath (Join-Path $SourceTclRoot 'tk8.6'))) {
    throw 'Tcl/Tk build libraries were not found.'
}
New-Item -ItemType Directory -Path $BuildTclRoot -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $SourceTclRoot 'tcl8.6') -Destination $BuildTclRoot -Recurse -Force
Copy-Item -LiteralPath (Join-Path $SourceTclRoot 'tk8.6') -Destination $BuildTclRoot -Recurse -Force

$PreviousTclLibrary = $env:TCL_LIBRARY
$PreviousTkLibrary = $env:TK_LIBRARY
$env:TCL_LIBRARY = Join-Path $BuildTclRoot 'tcl8.6'
$env:TK_LIBRARY = Join-Path $BuildTclRoot 'tk8.6'

Push-Location $ProjectDir
try {
    & $BuildPython -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --splash 'assets\splash.png' `
        --splash-center 'active' `
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
    if ($null -eq $PreviousTclLibrary) { Remove-Item Env:TCL_LIBRARY -ErrorAction SilentlyContinue } else { $env:TCL_LIBRARY = $PreviousTclLibrary }
    if ($null -eq $PreviousTkLibrary) { Remove-Item Env:TK_LIBRARY -ErrorAction SilentlyContinue } else { $env:TK_LIBRARY = $PreviousTkLibrary }
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
$SplashIncluded = (
    $ArchiveListing -match '(?m)^ Splash-00\.res$' -and
    $ArchiveListing -match '(?m)^ pyi_splash$'
)
if (-not $SplashIncluded) {
    throw 'EXE verification failed: startup splash was not bundled.'
}

Write-Host "EXE build completed: $ExePath"
