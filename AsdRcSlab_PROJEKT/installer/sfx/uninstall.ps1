$ErrorActionPreference='SilentlyContinue'
$appPlugins = Join-Path $env:APPDATA 'Autodesk\ApplicationPlugins'
Remove-Item (Join-Path $appPlugins 'AsdRcSlab.bundle') -Recurse -Force
$base='HKCU:\Software\Autodesk\AutoCAD'
if (Test-Path $base) {
  Get-ChildItem $base | ForEach-Object {
    Get-ChildItem $_.PSPath | ForEach-Object {
      $k = Join-Path (Join-Path $_.PSPath 'Applications') 'AsdRcSlab'
      if (Test-Path $k) { Remove-Item $k -Recurse -Force }
    }
  }
}
Add-Type -AssemblyName System.Windows.Forms
[void][System.Windows.Forms.MessageBox]::Show("AsdRcSlab uninstalled. Restart ASD.","AsdRcSlab uninstaller")
