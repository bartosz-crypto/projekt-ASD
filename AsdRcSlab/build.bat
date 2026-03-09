@echo off
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" "%~dp0AsdRcSlab.csproj" /p:Configuration=Release /p:OutputPath="%~dp0bin\Publish\\" /t:Build /v:minimal
