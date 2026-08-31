param(
    [switch]$ShowNotificationSample
)

$ErrorActionPreference = 'Stop'
$BundledPython = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectPython = Join-Path $ProjectDir '.venv-build\Scripts\python.exe'
$PythonExe = if (Test-Path -LiteralPath $ProjectPython) { $ProjectPython } else { $BundledPython }

if (Test-Path -LiteralPath $PythonExe) {
    # Tcl/Tk may fail when its library path contains non-ASCII characters.
    # Keep a runtime copy in an ASCII-only path.
    $BundledRoot = Split-Path -Parent $BundledPython
    $TkRuntime = 'C:\Users\Public\Documents\ESTsoft\CreatorTemp\dday-manager-tcl'
    $TclTarget = Join-Path $TkRuntime 'tcl8.6'
    $TkTarget = Join-Path $TkRuntime 'tk8.6'
    if (-not (Test-Path -LiteralPath (Join-Path $TclTarget 'init.tcl'))) {
        New-Item -ItemType Directory -Force -Path $TkRuntime | Out-Null
        Copy-Item -Recurse -Force -LiteralPath (Join-Path $BundledRoot 'tcl\tcl8.6') -Destination $TkRuntime
        Copy-Item -Recurse -Force -LiteralPath (Join-Path $BundledRoot 'tcl\tk8.6') -Destination $TkRuntime
    }
    $env:TCL_LIBRARY = $TclTarget
    $env:TK_LIBRARY = $TkTarget
    if ($ShowNotificationSample) {
        Push-Location $ProjectDir
        try {
            & $PythonExe -c "from notifications import show_windows_notification; show_windows_notification('내 일정 관리하기 - 알림 샘플', '오늘 마감 2건 · 3일 이내 4건 · 기한 초과 1건이 있습니다. 앱에서 확인하세요.')"
        } finally {
            Pop-Location
        }
    } else {
        & $PythonExe (Join-Path $ProjectDir 'app.py')
    }
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    python (Join-Path $ProjectDir 'app.py')
} else {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show('Python runtime was not found. See README.md.', 'D-Day Manager')
}
