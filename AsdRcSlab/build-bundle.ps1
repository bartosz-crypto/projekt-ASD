# build-bundle.ps1 - tworzy paczkę .bundle do dystrybucji w zespole (AutoCAD)
param(
    [string]$Configuration = "Release",
    [string]$Version = "2026.05"
)

$ErrorActionPreference = "Stop"

# === Ścieżki ===
$bundleSrc  = "$PSScriptRoot\bundle-template\AsdRcSlab.bundle"
$bundleOut  = "$PSScriptRoot\dist\AsdRcSlab.bundle"
$zipOut     = "$PSScriptRoot\dist\AsdRcSlab-$Version.zip"

# === Czyszczenie dist/ ===
if (Test-Path "$PSScriptRoot\dist") { Remove-Item "$PSScriptRoot\dist" -Recurse -Force }
New-Item -ItemType Directory -Path "$PSScriptRoot\dist" -Force | Out-Null

# === Build ===
Write-Host "Building..." -ForegroundColor Cyan
dotnet build "$PSScriptRoot\AsdRcSlab.csproj" -c $Configuration
if ($LASTEXITCODE -ne 0) { throw "Build failed" }

# === Kopiowanie struktury bundle ===
Write-Host "Copying bundle template..." -ForegroundColor Cyan
Copy-Item -Path $bundleSrc -Destination $bundleOut -Recurse

# === Kopiowanie DLL + zależności do Contents/ ===
# csproj ma <OutputPath>bin\Publish\</OutputPath> (bez podfolderu konfiguracji)
$binPath      = "$PSScriptRoot\bin\Publish"
$contentsPath = "$bundleOut\Contents"

if (-not (Test-Path $binPath)) { throw "Nie znaleziono output folderu: $binPath" }

Write-Host "Copying DLLs from $binPath..." -ForegroundColor Cyan

# Wyklucz: AutoCAD assemblies (są w procesie AutoCAD), artefakty build, pliki nie-dll
$skipList = @(
    "acmgd.dll", "acdbmgd.dll", "accoremgd.dll",
    "AcCoreMgd.dll", "AcDbMgd.dll", "AcMgd.dll",
    "AdWindows.dll", "AcWindows.dll",
    "AsdRcSlab_new.dll"      # artefakt testowy
)

Get-ChildItem "$binPath\*.dll" | Where-Object {
    $skipList -notcontains $_.Name
} | ForEach-Object {
    Write-Host "  + $($_.Name)"
    Copy-Item $_.FullName "$contentsPath\"
}

# === Aktualizacja wersji w PackageContents.xml ===
$pkgXml  = "$bundleOut\PackageContents.xml"
$content = Get-Content $pkgXml -Raw
$content = $content -replace 'FriendlyVersion="[^"]+"', "FriendlyVersion=`"$Version`""
Set-Content -Path $pkgXml -Value $content -Encoding utf8

# === ZIP do dystrybucji ===
Write-Host "Creating ZIP..." -ForegroundColor Cyan
Compress-Archive -Path $bundleOut -DestinationPath $zipOut

$zipSizeMB = [math]::Round((Get-Item $zipOut).Length / 1MB, 1)
$dlls = (Get-ChildItem "$contentsPath\*.dll").Count

Write-Host ""
Write-Host "==============================================" -ForegroundColor Green
Write-Host " Bundle gotowy:" -ForegroundColor Green
Write-Host "   $bundleOut  ($dlls DLL)" -ForegroundColor Yellow
Write-Host "   $zipOut  ($zipSizeMB MB)" -ForegroundColor Yellow
Write-Host "==============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Instalacja (AutoCAD):" -ForegroundColor Cyan
Write-Host '  Skopiuj AsdRcSlab.bundle do:' -ForegroundColor White
Write-Host '  %APPDATA%\Autodesk\ApplicationPlugins\' -ForegroundColor Yellow
