# Buduje SFX installer dist\AsdRcSlab_Setup_2026.05.exe z dist\AsdRcSlab.bundle.
# KLUCZOWE (p144): czysci staged folder ORAZ usuwa stary exe przed budowa —
# WinRAR 'a' DOPISUJE do istniejacego archiwum, wiec stary (zagniezdzony) exe
# zostawialby duplikaty -> AsdRcSlab.bundle\AsdRcSlab.bundle. Zawsze od zera.
$ErrorActionPreference = 'Stop'
$root   = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent   # ...\AsdRcSlab_PROJEKT
$sfx    = $PSScriptRoot
$src    = Join-Path $root 'dist\AsdRcSlab.bundle'
$staged = Join-Path $sfx  'AsdRcSlab.bundle'
$exe    = Join-Path $root 'dist\AsdRcSlab_Setup_2026.05.exe'
$winrar = 'C:\Program Files\WinRAR\WinRAR.exe'

if (-not (Test-Path $src)) { throw "Source bundle missing: $src (run build-bundle.ps1 first)" }

# 1) czysty staging (kopiuj FOLDER, ale do uprzednio usunietego miejsca -> brak zagniezdzenia)
if (Test-Path $staged) { Remove-Item $staged -Recurse -Force }
Copy-Item $src $staged -Recurse -Force
if (Test-Path (Join-Path $staged 'AsdRcSlab.bundle')) { throw "Nested staged bundle detected" }
if (-not (Test-Path (Join-Path $staged 'Contents\AsdRcSlab.dll'))) { throw "Staged DLL missing in Contents" }

# 2) usun stary exe (WinRAR 'a' dopisuje -> trzeba zaczac od zera)
if (Test-Path $exe) { Remove-Item $exe -Force }

# 3) zbuduj SFX
Push-Location $sfx
try {
  $p = Start-Process -Wait -PassThru -FilePath $winrar -ArgumentList `
    'a','-r','-sfx','-z"sfx.conf"','-ep1', "`"$exe`"", '"AsdRcSlab.bundle"','"setup.ps1"','"uninstall.ps1"'
  if ($p.ExitCode -ne 0) { throw "WinRAR failed exit=$($p.ExitCode)" }
} finally { Pop-Location }

Write-Output ("SFX built: " + $exe)
Get-Item $exe | Select-Object Name, Length, LastWriteTime
