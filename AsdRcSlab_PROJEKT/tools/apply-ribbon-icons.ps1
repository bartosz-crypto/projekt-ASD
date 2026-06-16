# apply-ribbon-icons.ps1 — podmienia ZAWARTOSC ikon przyciskow wstazki na pliki
# z folderu uzytkownika. Zachowuje nazwy docelowe (asd-<cmd>-16/-32.png), wiec
# RibbonBuilder/CuixBuilder nie wymagaja zmian sciezek.
#
# Zrodlo:  C:\Users\Dell\Desktop\ClaudeCode\ikony  (png/bmp/jpg/ico)
# Cel:     <project>\icons\<baza>-16.png  oraz  -32.png  (nadpisuje)

Add-Type -AssemblyName System.Drawing

$srcDir   = 'C:\Users\Dell\Desktop\ClaudeCode\ikony'
$iconsDir = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\icons'))

# 12 baz = 12 komend wstazki (kolejnosc jak na ribbonie).
$bases = @(
  'asd-gai','asd-rcn',                 # TITLE BLOCK
  'asd-pxie','asd-paa','asd-phr','asd-phv',  # PH CONDITIONS
  'asd-imr',                            # REINFORCEMENT MAPS
  'asd-bbc','asd-bbs-write',            # BBS
  'asd-xas','asd-sdc','asd-prg'         # TOOLS
)

$srcFiles = @(Get-ChildItem $srcDir -File | Sort-Object Name)
Write-Host ("=== Source folder: {0} ({1} files) ===" -f $srcDir, $srcFiles.Count)
$srcFiles | ForEach-Object { Write-Host ("  {0}" -f $_.Name) }
if ($srcFiles.Count -eq 0) { throw "Brak plikow zrodlowych w $srcDir" }

function Load-Image([string]$path) {
  if ([IO.Path]::GetExtension($path).ToLower() -eq '.ico') {
    $ico = New-Object System.Drawing.Icon($path)
    $bmp = $ico.ToBitmap()
    $ico.Dispose()
    return $bmp
  }
  return [System.Drawing.Image]::FromFile($path)
}

function Save-Scaled($img, [int]$size, [string]$outPath) {
  $bmp = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.Clear([System.Drawing.Color]::Transparent)
  $g.DrawImage($img, 0, 0, $size, $size)
  $g.Dispose()
  $bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose()
}

Write-Host "`n=== Mapping (command <- source icon) ==="
for ($i = 0; $i -lt $bases.Count; $i++) {
  $src = $srcFiles[$i % $srcFiles.Count]   # cyklicznie jesli mniej niz 12
  $img = Load-Image $src.FullName
  foreach ($sz in 16, 32) {
    $out = Join-Path $iconsDir ("{0}-{1}.png" -f $bases[$i], $sz)
    Save-Scaled $img $sz $out
  }
  $img.Dispose()
  Write-Host ("  {0,-16} <- {1}" -f $bases[$i], $src.Name)
}
Write-Host ("`nDone. Updated {0} icon bases (x2 sizes) in {1}" -f $bases.Count, $iconsDir) -ForegroundColor Green
