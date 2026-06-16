# gen-icons.ps1 — generuje proste ikony przyciskow wstazki ASD RC SLAB.
# Dla kazdej komendy: PNG 32x32 (large) i 16x16 (small) z alpha:
#   zaokraglony kwadrat w kolorze panelu + 2-literowe biale, pogrubione inicjaly.
# Pliki: <project>\icons\<iconName>-<size>.png  (build-bundle kopiuje icons\ -> Contents\Icons).

Add-Type -AssemblyName System.Drawing

$outDir = Join-Path $PSScriptRoot "..\icons"
$outDir = [System.IO.Path]::GetFullPath($outDir)
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

# iconName ; inicjaly ; kolor panelu (#RRGGBB)
$icons = @(
  @{ name="asd-gai";       text="GA"; color="#1565C0" }  # TITLE BLOCK
  @{ name="asd-rcn";       text="SN"; color="#1565C0" }
  @{ name="asd-pxie";      text="LP"; color="#2E7D32" }  # PH CONDITIONS
  @{ name="asd-paa";       text="PH"; color="#2E7D32" }
  @{ name="asd-phr";       text="PR"; color="#2E7D32" }
  @{ name="asd-phv";       text="PV"; color="#2E7D32" }
  @{ name="asd-imr";       text="IM"; color="#EF6C00" }  # REINFORCEMENT MAPS
  @{ name="asd-bbc";       text="BC"; color="#6A1B9A" }  # BBS
  @{ name="asd-bbs-write"; text="BW"; color="#6A1B9A" }
  @{ name="asd-xas";       text="XA"; color="#00838F" }  # TOOLS
  @{ name="asd-sdc";       text="SD"; color="#00838F" }
  @{ name="asd-prg";       text="PG"; color="#00838F" }
)

function ColorFromHex([string]$hex) {
  $hex = $hex.TrimStart('#')
  $r = [Convert]::ToInt32($hex.Substring(0,2),16)
  $g = [Convert]::ToInt32($hex.Substring(2,2),16)
  $b = [Convert]::ToInt32($hex.Substring(4,2),16)
  return [System.Drawing.Color]::FromArgb(255,$r,$g,$b)
}

function New-RoundedPath([int]$size,[int]$radius) {
  $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
  $m = 1                     # marginez, zeby alpha krawedzie byla gladka
  $x = $m; $y = $m; $w = $size - 2*$m; $h = $size - 2*$m; $d = 2*$radius
  $gp.AddArc($x, $y, $d, $d, 180, 90)
  $gp.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
  $gp.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
  $gp.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
  $gp.CloseFigure()
  return $gp
}

function Make-Icon([string]$path,[int]$size,[string]$text,[System.Drawing.Color]$bg) {
  $bmp = New-Object System.Drawing.Bitmap($size,$size,[System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.TextRenderingHint  = [System.Drawing.Text.TextRenderingHint]::AntiAlias
  $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.Clear([System.Drawing.Color]::Transparent)

  $radius = [int]([Math]::Max(2, $size * 0.22))
  $path2  = New-RoundedPath $size $radius
  $brush  = New-Object System.Drawing.SolidBrush($bg)
  $g.FillPath($brush, $path2)

  $fontSize = [single]($size * 0.40)
  $font = New-Object System.Drawing.Font("Segoe UI", $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
  $sf = New-Object System.Drawing.StringFormat
  $sf.Alignment     = [System.Drawing.StringAlignment]::Center
  $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
  $rect = New-Object System.Drawing.RectangleF(0, 0, $size, $size)
  $white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
  $g.DrawString($text, $font, $white, $rect, $sf)

  $g.Dispose()
  $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose()
  $brush.Dispose(); $white.Dispose(); $font.Dispose(); $sf.Dispose(); $path2.Dispose()
}

$count = 0
foreach ($ic in $icons) {
  $bg = ColorFromHex $ic.color
  foreach ($sz in 32,16) {
    $file = Join-Path $outDir ("{0}-{1}.png" -f $ic.name, $sz)
    Make-Icon $file $sz $ic.text $bg
    $count++
    Write-Host ("  + {0}" -f (Split-Path $file -Leaf))
  }
}
Write-Host ("Generated {0} icon files in {1}" -f $count, $outDir) -ForegroundColor Green
