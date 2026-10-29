param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'splash.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function ConvertFrom-CodePoints {
    param([int[]]$CodePoints)
    return -join ($CodePoints | ForEach-Object { [char]$_ })
}

$subtitleText = ConvertFrom-CodePoints @(0xC5C5, 0xBB34, 0xC640, 0x20, 0xC77C, 0xC815, 0xC744, 0x20, 0xC900, 0xBE44, 0xD558, 0xACE0, 0x20, 0xC788, 0xC2B5, 0xB2C8, 0xB2E4)
$waitText = ConvertFrom-CodePoints @(0xC7A0, 0xC2DC, 0xB9CC, 0x20, 0xAE30, 0xB2E4, 0xB824, 0x20, 0xC8FC, 0xC138, 0xC694)
$ownerName = ConvertFrom-CodePoints @(0xC774, 0xC8FC, 0xC0BC)
$copyrightText = ([char]0x00A9) + ' 2026 ' + $ownerName

$width = 720
$height = 400
$bitmap = New-Object System.Drawing.Bitmap($width, $height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$icon = $null
$privateFonts = $null
$titleFont = $null
$subtitleFont = $null
$smallFont = $null

try {
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $graphics.Clear([System.Drawing.ColorTranslator]::FromHtml('#F4FBEF'))

    $leftBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#E4F5DA'))
    $graphics.FillRectangle($leftBrush, 0, 0, 260, $height)
    $leftBrush.Dispose()

    $borderPen = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml('#BFD8B8'), 2)
    $graphics.DrawRectangle($borderPen, 1, 1, $width - 3, $height - 3)
    $borderPen.Dispose()

    $iconPath = Join-Path $PSScriptRoot 'app_icon.png'
    $icon = [System.Drawing.Image]::FromFile($iconPath)
    $graphics.DrawImage($icon, 45, 88, 170, 170)

    $privateFonts = New-Object System.Drawing.Text.PrivateFontCollection
    $privateFonts.AddFontFile((Join-Path $PSScriptRoot 'fonts\Fredoka-Bold.ttf'))
    $titleFamily = $privateFonts.Families[0]
    $titleFont = New-Object System.Drawing.Font($titleFamily, 36, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point)
    $subtitleFont = New-Object System.Drawing.Font('Malgun Gothic', 15, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
    $smallFont = New-Object System.Drawing.Font('Malgun Gothic', 10, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)

    $titleBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#24324A'))
    $textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#415348'))
    $mutedBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#758477'))
    $graphics.DrawString('Task It Easy', $titleFont, $titleBrush, 305, 102)
    $graphics.DrawString($subtitleText, $subtitleFont, $textBrush, 308, 184)
    $graphics.DrawString($waitText, $smallFont, $mutedBrush, 310, 226)

    foreach ($x in @(312, 334, 356)) {
        $dotBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#43A653'))
        $graphics.FillEllipse($dotBrush, $x, 266, 10, 10)
        $dotBrush.Dispose()
    }

    $graphics.DrawString($copyrightText, $smallFont, $mutedBrush, 555, 354)
    $titleBrush.Dispose()
    $textBrush.Dispose()
    $mutedBrush.Dispose()

    $outputDirectory = Split-Path -Parent $OutputPath
    if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
    }
    $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Host "Splash image created: $OutputPath"
} finally {
    if ($smallFont) { $smallFont.Dispose() }
    if ($subtitleFont) { $subtitleFont.Dispose() }
    if ($titleFont) { $titleFont.Dispose() }
    if ($privateFonts) { $privateFonts.Dispose() }
    if ($icon) { $icon.Dispose() }
    $graphics.Dispose()
    $bitmap.Dispose()
}
