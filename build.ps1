# ═══════════════════════════════════════════════════════════════
#  Build Script: Combines modular planner-dev files into a
#  single standalone HTML file.
# ═══════════════════════════════════════════════════════════════

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$outFile = Join-Path $scriptDir '..\H-Customer_Performance_Planner_BUILD.html'

Write-Host '[BUILD] Starting build...'

# Read CSS
$css = Get-Content -Raw 'src\css\style.css' -Encoding UTF8

# Read all JS files in correct load order
$jsOrder = @(
  'state.js',
  'render.js',
  'interactions.js',
  'items.js',
  'holidays.js',
  'links.js',
  'alarms.js',
  'csv.js',
  'ui.js'
)
$allJs = ''
foreach ($f in $jsOrder) {
  $content = Get-Content -Raw "src\js\$f" -Encoding UTF8
  $allJs += "`n// === $f ===`n" + $content + "`n"
}

# Read logo as base64
$logoPath = (Resolve-Path 'src\img\logo.jpg').Path
$logoBytes = [System.IO.File]::ReadAllBytes($logoPath)
$logoB64 = [System.Convert]::ToBase64String($logoBytes)
$logoDataUri = 'data:image/jpeg;base64,' + $logoB64

# Read index.html
$html = Get-Content -Raw 'index.html' -Encoding UTF8

# Replace CSS link with inline style
$html = $html -replace '<link rel="stylesheet" href="src/css/style.css" />', "<style>`n$css`n  </style>"

# Replace logo src
$html = $html -replace 'src="src/img/logo.jpg"', "src=`"$logoDataUri`""

# Replace script tags with single inline script block
$scriptPattern = '(?s)<!-- JS Modules.*?</script>'
$html = $html -replace $scriptPattern, "<script>`n$allJs`n  </script>"

# Write output with UTF-8 BOM
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($outFile, $html, $utf8Bom)

$size = [math]::Round((Get-Item $outFile).Length / 1KB)
Write-Host "[BUILD] Output: $outFile"
Write-Host "[BUILD] Size: $size KB"
Write-Host '[BUILD] Done!'
