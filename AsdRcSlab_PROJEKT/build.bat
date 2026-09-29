@echo off
REM p165: projekt jest SDK-style (Microsoft.NET.Sdk) - stary MSBuild v4.0 go nie obsluguje.
REM Output: bin\Publish\AsdRcSlab.dll (OutputPath z csproj).
dotnet build "%~dp0AsdRcSlab.csproj" -c Release -v minimal
if errorlevel 1 (echo. & echo BUILD FAILED & pause & exit /b 1)
echo. & echo BUILD OK: %~dp0bin\Publish\AsdRcSlab.dll
pause
