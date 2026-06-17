$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
try {
  # 0) NIE instaluj, gdy ASD/AutoCAD dziala (inaczej DLL zablokowana -> zagniezdzenie)
  if (Get-Process -Name acad -ErrorAction SilentlyContinue) {
    [void][System.Windows.Forms.MessageBox]::Show("Close AutoCAD/ASD completely, then run the installer again.","AsdRcSlab installer")
    return
  }
  $appPlugins = Join-Path $env:APPDATA 'Autodesk\ApplicationPlugins'
  $dst = Join-Path $appPlugins 'AsdRcSlab.bundle'
  # 1) wyczysc POPRZEDNI install w calosci (w tym ewentualne zagniezdzenie)
  if (Test-Path $dst) { Remove-Item $dst -Recurse -Force }
  New-Item $dst -ItemType Directory -Force | Out-Null
  # 2) kopiuj ZAWARTOSC staged bundla (z gwiazdka!) do $dst - NIE sam folder (brak zagniezdzenia)
  Copy-Item (Join-Path $PSScriptRoot 'AsdRcSlab.bundle\*') $dst -Recurse -Force
  # 3) walidacja: DLL musi byc w Contents\ (tam wskazuje LOADER)
  $loader = Join-Path $dst 'Contents\AsdRcSlab.dll'
  if (-not (Test-Path $loader)) { throw "DLL missing after copy: $loader" }
  if (Test-Path (Join-Path $dst 'AsdRcSlab.bundle')) { throw "Nested bundle detected after copy" }
  # 4) rejestr autoload - enumeracja profili (jak dotad)
  function Write-AppKey($appsKey){
    $k=Join-Path $appsKey 'AsdRcSlab'; New-Item -Path $k -Force | Out-Null
    New-ItemProperty $k -Name DESCRIPTION -Value 'AsdRcSlab plugin commands and ribbon' -PropertyType String -Force | Out-Null
    New-ItemProperty $k -Name LOADCTRLS -Value 2 -PropertyType DWord -Force | Out-Null
    New-ItemProperty $k -Name LOADER   -Value $loader -PropertyType String -Force | Out-Null
    New-ItemProperty $k -Name MANAGED  -Value 1 -PropertyType DWord -Force | Out-Null
  }
  $cnt=0; $base='HKCU:\Software\Autodesk\AutoCAD'
  if (Test-Path $base){ Get-ChildItem $base -EA SilentlyContinue | %{ Get-ChildItem $_.PSPath -EA SilentlyContinue | %{ $a=Join-Path $_.PSPath 'Applications'; if(Test-Path $a){ Write-AppKey $a; $cnt++ } } } }
  $bk='HKCU:\Software\Autodesk\AutoCAD\R20.0\ACAD-E030:409\Applications'
  if(-not (Test-Path (Join-Path $bk 'AsdRcSlab'))){ New-Item $bk -Force | Out-Null; Write-AppKey $bk; $cnt++ }
  # 5) uninstaller
  $undir = Join-Path $appPlugins 'AsdRcSlab_uninstall'; New-Item $undir -ItemType Directory -Force | Out-Null
  Copy-Item (Join-Path $PSScriptRoot 'uninstall.ps1') (Join-Path $undir 'uninstall.ps1') -Force
  [void][System.Windows.Forms.MessageBox]::Show("AsdRcSlab installed (profiles: $cnt).`r`nStart ASD and run ASD-PRG - the command line must show the build stamp.","AsdRcSlab installer")
} catch {
  [void][System.Windows.Forms.MessageBox]::Show("Install failed: $($_.Exception.Message)","AsdRcSlab installer")
}
