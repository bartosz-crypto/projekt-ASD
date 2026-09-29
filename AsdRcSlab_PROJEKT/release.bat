@echo off
REM p165: pelny release jednym dwuklikiem: build + bundle + instalator SFX.
REM Pliki .ps1 odpalone z cmd/dwuklikiem otwieraja sie w Notatniku - dlatego
REM wolamy je jawnie przez powershell z ExecutionPolicy Bypass.
REM Wynik: dist\AsdRcSlab.bundle + dist\AsdRcSlab_Setup_<wersja>.exe
cd /d "%~dp0"

tasklist /FI "IMAGENAME eq acad.exe" | find /I "acad.exe" >nul
if not errorlevel 1 (
  echo Zamknij AutoCAD/ASD przed buildem - trzyma zablokowany DLL.
  pause
  exit /b 1
)

echo === 1/2 build-bundle.ps1 ===
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-bundle.ps1"
if errorlevel 1 (echo. & echo BUNDLE FAILED & pause & exit /b 1)

echo.
echo === 2/2 installer\sfx\build-sfx.ps1 ===
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer\sfx\build-sfx.ps1"
if errorlevel 1 (echo. & echo SFX FAILED & pause & exit /b 1)

echo.
echo RELEASE OK - instalator w folderze dist\
pause
