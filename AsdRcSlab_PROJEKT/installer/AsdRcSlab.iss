; ============================================================================
;  AsdRcSlab.iss  — instalator Inno Setup wtyczki ASD RC SLAB
;  Replika metody kolegi (Setup 4.4):
;    - kopiuje bundle do %APPDATA%\Autodesk\ApplicationPlugins\AsdRcSlab.bundle
;    - pisze HKCU ...\Applications\AsdRcSlab (LOADER, MANAGED=1, LOADCTRLS=2,
;      DESCRIPTION) pod KAŻDYM wykrytym produktem AutoCAD/ASD (early-load)
;  Per-user (bez admina). v1 BEZ podpisu cert.
;  Kompilacja:  "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\AsdRcSlab.iss
; ============================================================================

#define MyAppName "AsdRcSlab"
#define MyAppVersion "2026.05"
#define MyDll "{userappdata}\Autodesk\ApplicationPlugins\AsdRcSlab.bundle\Contents\AsdRcSlab.dll"
#define BaseKey "Software\Autodesk\AutoCAD\R20.0\ACAD-E030:409\Applications\AsdRcSlab"

[Setup]
AppId={{4E200B1C-B504-4C88-9275-B1253BD4F1F7}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=NA Engineering sp. z o.o.
AppPublisherURL=https://naengineering.uk
PrivilegesRequired=lowest
DefaultDirName={userappdata}\Autodesk\ApplicationPlugins\AsdRcSlab_uninstall
DisableDirPage=yes
DisableProgramGroupPage=yes
Uninstallable=yes
OutputDir=..\dist
OutputBaseFilename=AsdRcSlab_Setup_{#MyAppVersion}
SetupLogging=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"

[Files]
; Cały bundle rekurencyjnie -> ApplicationPlugins\AsdRcSlab.bundle
Source: "..\dist\AsdRcSlab.bundle\*"; \
  DestDir: "{userappdata}\Autodesk\ApplicationPlugins\AsdRcSlab.bundle"; \
  Flags: recursesubdirs createallsubdirs ignoreversion

[UninstallDelete]
; Przy deinstalacji usuń cały folder bundla
Type: filesandordirs; Name: "{userappdata}\Autodesk\ApplicationPlugins\AsdRcSlab.bundle"

[Registry]
; Pewny baseline: potwierdzony produkt usera (ASD 2015 EN = R20.0 / ACAD-E030:409).
; Dodatkowe produkty dopisuje [Code] (enumeracja). Ten klucz znika przy uninstall.
Root: HKCU; Subkey: "{#BaseKey}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "{#BaseKey}"; ValueType: string; ValueName: "DESCRIPTION"; ValueData: "AsdRcSlab plugin commands and ribbon"
Root: HKCU; Subkey: "{#BaseKey}"; ValueType: dword;  ValueName: "LOADCTRLS"; ValueData: "2"
Root: HKCU; Subkey: "{#BaseKey}"; ValueType: string; ValueName: "LOADER"; ValueData: "{#MyDll}"
Root: HKCU; Subkey: "{#BaseKey}"; ValueType: dword;  ValueName: "MANAGED"; ValueData: "1"

[Code]
const
  AUTOCAD_ROOT = 'Software\Autodesk\AutoCAD';

// Zapisuje 4 wartosci demand-load pod danym kluczem Applications\AsdRcSlab.
procedure WriteAppKey(const Base: String);
begin
  RegWriteStringValue(HKEY_CURRENT_USER, Base, 'DESCRIPTION', 'AsdRcSlab plugin commands and ribbon');
  RegWriteDWordValue (HKEY_CURRENT_USER, Base, 'LOADCTRLS', 2);
  RegWriteStringValue(HKEY_CURRENT_USER, Base, 'LOADER', ExpandConstant('{#MyDll}'));
  RegWriteDWordValue (HKEY_CURRENT_USER, Base, 'MANAGED', 1);
end;

// (a) AutoCAD/ASD zainstalowany?  (b) czy nie jest uruchomiony?
function InitializeSetup(): Boolean;
var
  Vers: TArrayOfString;
  ResultCode: Integer;
begin
  Result := True;

  if (not RegGetSubkeyNames(HKEY_CURRENT_USER, AUTOCAD_ROOT, Vers)) or (GetArrayLength(Vers) = 0) then
  begin
    MsgBox('Nie wykryto AutoCAD / ASD na tym koncie uzytkownika' + #13#10 +
           '(brak HKCU\' + AUTOCAD_ROOT + ').' + #13#10 + #13#10 +
           'Zainstaluj/uruchom raz AutoCAD lub ASD, potem ponow instalacje.',
           mbError, MB_OK);
    Result := False;
    exit;
  end;

  // tasklist + find: ResultCode=0 gdy acad.exe znaleziony.
  if Exec(ExpandConstant('{cmd}'),
          '/C tasklist /FI "IMAGENAME eq acad.exe" | find /I "acad.exe"',
          '', SW_HIDE, ewWaitUntilTerminated, ResultCode) and (ResultCode = 0) then
  begin
    if MsgBox('AutoCAD / ASD jest uruchomiony.' + #13#10 +
              'Zamknij go przed instalacja (inaczej DLL moze byc zablokowany).' + #13#10 + #13#10 +
              'Kontynuowac mimo to?', mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;

// (c) "u wszystkich": dopisz klucz pod KAZDYM wykrytym produktem AutoCAD/ASD.
procedure CurStepChanged(CurStep: TSetupStep);
var
  Vers, Prods: TArrayOfString;
  i, j: Integer;
  VerPath, ProdPath: String;
begin
  if CurStep <> ssPostInstall then exit;

  if not RegGetSubkeyNames(HKEY_CURRENT_USER, AUTOCAD_ROOT, Vers) then exit;

  for i := 0 to GetArrayLength(Vers) - 1 do
  begin
    VerPath := AUTOCAD_ROOT + '\' + Vers[i];
    if not RegGetSubkeyNames(HKEY_CURRENT_USER, VerPath, Prods) then continue;

    for j := 0 to GetArrayLength(Prods) - 1 do
    begin
      ProdPath := VerPath + '\' + Prods[j];
      // tylko realne profile produktu (maja podklucz Applications)
      if RegKeyExists(HKEY_CURRENT_USER, ProdPath + '\Applications') then
        WriteAppKey(ProdPath + '\Applications\AsdRcSlab');
    end;
  end;
end;
