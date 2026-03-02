using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AsdRcSlab
{
    public static class PunchingParser
    {
        public static List<string> GetSheetNames(string xlsxPath)
        {
            using (var pkg = new ExcelPackage(new FileInfo(xlsxPath)))
                return pkg.Workbook.Worksheets.Select(w => w.Name).ToList();
        }

        public static List<PileData> Parse(string xlsxPath, string sheetName, out string parseLog)
        {
            var piles = new List<PileData>();
            var log   = new System.Text.StringBuilder();

            using (var pkg = new ExcelPackage(new FileInfo(xlsxPath)))
            {
                ExcelWorksheet ws = pkg.Workbook.Worksheets[sheetName];
                if (ws == null)
                {
                    log.AppendLine($"Nie znaleziono arkusza '{sheetName}'.");
                    parseLog = log.ToString();
                    return piles;
                }

                int lastRow = ws.Dimension?.End.Row    ?? 1;
                int lastCol = ws.Dimension?.End.Column ?? 80;

                log.AppendLine($"Arkusz: '{sheetName}', wiersze: {lastRow}, kolumny: {lastCol}");

                // ── Krok 1: Znajdź kolumnę ACTION (ADD H.../NO ACTION) ────────────────
                // Szuka ostatniej kolumny zawierającej wartości akcji w górnych wierszach
                int colAction = -1;

                // Skanuj wiersze danych (7–30) by znaleźć kolumnę z ADD H.../NO ACTION
                for (int c = lastCol; c >= 1; c--)
                {
                    bool found = false;
                    for (int r = 7; r <= Math.Min(30, lastRow); r++)
                    {
                        string v = ws.Cells[r, c].GetValue<string>()?.Trim() ?? "";
                        if (IsActionValue(v)) { found = true; break; }
                    }
                    if (found) { colAction = c; break; }
                }

                if (colAction < 0)
                {
                    log.AppendLine("Nie znaleziono kolumny ACTION (ADD H.../NO ACTION).");
                    // Dump pierwszych 8 wierszy col 1-5 dla diagnostyki
                    for (int r = 1; r <= Math.Min(8, lastRow); r++)
                    {
                        var vals = Enumerable.Range(1, 5)
                            .Select(c => ws.Cells[r, c].GetValue<string>() ?? "");
                        log.AppendLine($"  R{r}: {string.Join(" | ", vals)}");
                    }
                    parseLog = log.ToString();
                    return piles;
                }
                log.AppendLine($"Kolumna ACTION: {colAction}");

                // ── Krok 2: Zbierz etykiety lokalizacji z wierszy nagłówkowych (1–6) ──
                // Etykiety sekcji: "INTERNAL PILES", "CORNER PILES", "EDGE PILES"
                var locationLabels = new List<(int Row, string Location)>();
                for (int r = 1; r <= Math.Min(6, lastRow); r++)
                {
                    for (int c = 1; c <= Math.Min(lastCol, 60); c++)
                    {
                        string v = ws.Cells[r, c].GetValue<string>()?.Trim()?.ToUpperInvariant() ?? "";
                        string loc = "";
                        if (v.Contains("INTERNAL PILES")) loc = "INT";
                        else if (v.Contains("CORNER PILES")) loc = "CORNER";
                        else if (v.Contains("EDGE PILES"))   loc = "EDGE";

                        if (loc != "" && !locationLabels.Any(l => l.Location == loc))
                        {
                            locationLabels.Add((r, loc));
                            break;
                        }
                    }
                }

                log.AppendLine($"Znalezione etykiety lokalizacji: {string.Join(", ", locationLabels.Select(l => $"R{l.Row}={l.Location}"))}");

                // Etykiety posortowane malejąco (ostatni wiersz = najblizej danych = sekcja 1)
                var orderedLabels = locationLabels.OrderByDescending(l => l.Row).Select(l => l.Location).ToList();
                // orderedLabels[0] = sekcja 1, [1] = sekcja 2, itd.

                // Fallback: jeśli brak etykiet, przypisz domyślnie INT
                if (orderedLabels.Count == 0) orderedLabels.Add("INT");

                // ── Krok 3: Czytaj wiersze danych ────────────────────────────────────
                int sectionIndex = 0;

                for (int r = 7; r <= lastRow; r++)
                {
                    string col1 = ws.Cells[r, 1].GetValue<string>()?.Trim() ?? "";

                    // Separator sekcji ('o' lub 'O' lub pusta linia po kilku danych)
                    if (col1 == "o" || col1 == "O")
                    {
                        sectionIndex++;
                        continue;
                    }

                    // Pomijaj wiersze nagłówkowe w sekcjach (NIB REDUCTION, L [mm] itp.)
                    if (string.IsNullOrWhiteSpace(col1))    continue;
                    if (!IsPileId(col1))                    continue;

                    string pileId = col1;

                    // Lokalizacja z bieżącej sekcji
                    string loc = sectionIndex < orderedLabels.Count
                        ? orderedLabels[sectionIndex]
                        : orderedLabels[orderedLabels.Count - 1];

                    // Action
                    string action = ws.Cells[r, colAction].GetValue<string>()?.Trim() ?? "NO ACTION";

                    // Utylizacja z kolumny 9 (lub szukaj numerycznej wartości w okolicy)
                    double util = 0;
                    TryReadDouble(ws.Cells[r, 9].GetValue<string>(), out util);

                    piles.Add(new PileData
                    {
                        PileId         = pileId,
                        UtilPct        = util,
                        LocationType   = loc,
                        PunchingAction = action
                    });
                }

                log.AppendLine($"Wczytano {piles.Count} pali.");
                if (piles.Count > 0)
                {
                    var groups = piles.GroupBy(p => p.LocationType);
                    foreach (var g in groups)
                        log.AppendLine($"  {g.Key}: {g.Count()} pali — akcje: {string.Join(", ", g.Select(p => p.PunchingAction).Distinct())}");
                }
            }

            parseLog = log.ToString();
            return piles;
        }

        // ── Pomocnicze ──────────────────────────────────────────────────────────────

        private static bool IsActionValue(string v)
        {
            if (string.IsNullOrEmpty(v)) return false;
            string u = v.ToUpperInvariant();
            return u.StartsWith("ADD H") || u == "NO ACTION";
        }

        /// <summary>Czy wartość wygląda jak pile ID: 4-cyfrowa lub alfanumeryczna, nie jest nagłówkiem sekcji.</summary>
        private static bool IsPileId(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return false;
            string u = val.Trim().ToUpperInvariant();
            // Odrzuć znane nagłówki
            if (u.EndsWith("PILES"))    return false;
            if (u.Contains("REDUC"))    return false;
            if (u.Contains("SECTION")) return false;
            if (u == "]" || u == "O" || u == "I" || u == "V" || u == "C") return false;
            // Akceptuj liczby lub krótkie kody alfanumeryczne
            if (int.TryParse(u, out _)) return true;
            if (u.Length >= 2 && u.Length <= 8 && u.All(ch => char.IsLetterOrDigit(ch) || ch == '-'))
                return true;
            return false;
        }

        private static bool TryReadDouble(string s, out double val)
        {
            if (string.IsNullOrEmpty(s)) { val = 0; return false; }
            s = s.Replace(",", ".").Replace("%", "").Trim();
            return double.TryParse(s, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out val);
        }
    }
}
