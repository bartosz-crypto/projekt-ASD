$ErrorActionPreference = 'Stop'
try {
  $src = Join-Path $PSScriptRoot 'AsdRcSlab.bundle'
  $appPlugins = Join-Path $env:APPDATA 'Autodesk\ApplicationPlugins'
  $dst = Join-Path $appPlugins 'AsdRcSlab.bundle'
  if (Test-Path $dst) { Remove-Item $dst -Recurse -Force }
  New-Item $appPlugins -ItemType Directory -Force | Out-Null
  Copy-Item $src $dst -Recurse -Force
  $loader = Join-Path $dst 'Contents\AsdRcSlab.dll'
  function Write-AppKey($appsKey) {
    $k = Join-Path $appsKey 'AsdRcSlab'
    New-Item -Path $k -Force | Out-Null
    New-ItemProperty $k -Name DESCRIPTION -Value 'AsdRcSlab plugin commands and ribbon' -PropertyType String -Force | Out-Null
    New-ItemProperty $k -Name LOADCTRLS  -Value 2        -PropertyType DWord  -Force | Out-Null
    New-ItemProperty $k -Name LOADER     -Value $loader  -PropertyType String -Force | Out-Null
    New-ItemProperty $k -Name MANAGED    -Value 1        -PropertyType DWord  -Force | Out-Null
  }
  $count = 0
  $base = 'HKCU:\Software\Autodesk\AutoCAD'
  if (Test-Path $base) {
    Get-ChildItem $base -ErrorAction SilentlyContinue | ForEach-Object {
      Get-ChildItem $_.PSPath -ErrorAction SilentlyContinue | ForEach-Object {
        $apps = Join-Path $_.PSPath 'Applications'
        if (Test-Path $apps) { Write-AppKey $apps; $count++ }
      }
    }
  }
  $bk = 'HKCU:\Software\Autodesk\AutoCAD\R20.0\ACAD-E030:409\Applications'
  if (-not (Test-Path (Join-Path $bk 'AsdRcSlab'))) { New-Item $bk -Force | Out-Null; Write-AppKey $bk; $count++ }
  $undir = Join-Path $appPlugins 'AsdRcSlab_uninstall'
  New-Item $undir -ItemType Directory -Force | Out-Null
  Copy-Item (Join-Path $PSScriptRoot 'uninstall.ps1') (Join-Path $undir 'uninstall.ps1') -Force
  Add-Type -AssemblyName System.Windows.Forms
  [void][System.Windows.Forms.MessageBox]::Show("AsdRcSlab installed (autoload set for $count profile(s)).`r`nClose & restart ASD to load the plugin + ribbon.","AsdRcSlab installer")
} catch {
  Add-Type -AssemblyName System.Windows.Forms
  [void][System.Windows.Forms.MessageBox]::Show("Install failed: $($_.Exception.Message)","AsdRcSlab installer")
}
