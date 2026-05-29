# AsdRcSlab — Instalacja

## Wymagania

- AutoCAD 2015 lub nowszy (Structural Detailing preferred)
- Windows 64-bit
- .NET Framework 4.8

## Instalacja

1. Pobierz `AsdRcSlab-<wersja>.zip` z wewnętrznego repozytorium / dysku zespołu.
2. Rozpakuj. Wewnątrz znajdziesz folder `AsdRcSlab.bundle`.
3. **Zamknij AutoCAD** jeśli jest otwarty.
4. Skopiuj folder `AsdRcSlab.bundle` (cały folder, nie zawartość) do:
   ```
   %APPDATA%\Autodesk\ApplicationPlugins\
   ```
   (Wklej tę ścieżkę w pasek Eksploratora Windows.)
5. Uruchom AutoCAD. Plugin załaduje się automatycznie.
6. Sprawdź: na wstążce powinna pojawić się zakładka **ASD RC SLAB** z komendami.

## Aktualizacja

1. Zamknij AutoCAD.
2. Usuń stary folder `AsdRcSlab.bundle` z `%APPDATA%\Autodesk\ApplicationPlugins\`.
3. Wklej nowy folder z nowej paczki.
4. Uruchom AutoCAD.

## Deinstalacja

1. Zamknij AutoCAD.
2. Usuń `AsdRcSlab.bundle` z `%APPDATA%\Autodesk\ApplicationPlugins\`.

## Troubleshooting

**Komendy ASD-* nie działają / brak wstążki:**
- Sprawdź czy folder bundle jest w prawidłowej lokalizacji.
- W AutoCAD wpisz `APPAUTOLOAD`. Powinien pokazać `AsdRcSlab` jako loaded.
- Spróbuj `NETLOAD` ręcznie wskazując `AsdRcSlab.dll` z folderu `bundle\Contents\` — jeśli komendy działają po ręcznym NETLOAD, problem leży w `PackageContents.xml`.

**Plugin nie ładuje się po update AutoCAD:**
- Sprawdź czy nowsza wersja AutoCAD mieści się w zakresie `SeriesMin`–`SeriesMax` w `PackageContents.xml`.
- Zaktualizuj `SeriesMax` i przebuduj paczkę (`./build-bundle.ps1`).

**Błąd .NET przy ładowaniu:**
- Upewnij się że .NET Framework 4.8 jest zainstalowany (wbudowany w Windows 10 1903+).
