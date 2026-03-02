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

                // ── Krok 1: Znajdź kolumnę ACTION ────────────────────────────────────
                int colAction = FindActionColumn(ws, lastRow, lastCol, log);
                if (colAction < 0)
                {
                    log.AppendLine("Nie znaleziono kolumny ACTION (ADD H.../NO ACTION).");
                    DumpRows(ws, lastRow, log);
                    parseLog = log.ToString();
                    return piles;
                }
                log.AppendLine($"Kolumna ACTION: {colAction}");

                // ── Krok 2: Skanuj WSZYSTKIE wiersze w poszukiwaniu etykiet sekcji ──
                // Etykiety mogą być w nagłówkach (r 1-6) lub wewnątrz danych (r 7+)
                var sectionBoundaries = new List<(int Row, string Location)>();

                for (int r = 1; r <= lastRow; r++)
                {
                    string loc = DetectSectionLabel(ws, r, lastCol);
                    if (loc != null)
                        sectionBoundaries.Add((r, loc));
                }

                sectionBoundaries = sectionBoundaries.OrderBy(s => s.Row).ToList();
                log.AppendLine($"Wykryte sekcje: {string.Join(", ", sectionBoundaries.Select(s => $"R{s.Row}={s.Location}"))}");

                // ── Krok 3: Czytaj wiersze danych ────────────────────────────────────
                // Dwa tryby:
                // A) Etykiety w wierszach danych (row > 6) → granice po numerze wiersza
                // B) Etykiety tylko w nagłówkach (row <= 6) → granice po separatorze 'o'
                bool dataLabels = sectionBoundaries.Any(s => s.Row > 6);

                // Dla trybu B: kolejność sekcji = kolejność znalezienia (lewo-prawo)
                var headerOrder = sectionBoundaries
                    .Where(s => s.Row <= 6)
                    .Select(s => s.Location)
                    .ToList();
                if (headerOrder.Count == 0 && !dataLabels)
                    headerOrder.Add("INT"); // fallback

                string currentLocation = sectionBoundaries.Count > 0
                    ? sectionBoundaries[0].Location
                    : "INT";
                int secIdx      = 0;   // dla trybu A (row-based)
                int headerSecIdx = 0;  // dla trybu B ('o'-based)

                for (int r = 7; r <= lastRow; r++)
                {
                    // Tryb A: aktualizuj sekcję gdy przekroczyliśmy granicę wiersza
                    if (dataLabels)
                    {
                        while (secIdx + 1 < sectionBoundaries.Count &&
                               r >= sectionBoundaries[secIdx + 1].Row)
                        {
                            secIdx++;
                            currentLocation = sectionBoundaries[secIdx].Location;
                        }
                    }

                    string col1 = ws.Cells[r, 1].GetValue<string>()?.Trim() ?? "";

                    // Separator 'o'
                    if (col1 == "o" || col1 == "O")
                    {
                        // Tryb B: 'o' przesuwa sekcję
                        if (!dataLabels && headerOrder.Count > 0)
                        {
                            headerSecIdx = Math.Min(headerSecIdx + 1, headerOrder.Count - 1);
                            currentLocation = headerOrder[headerSecIdx];
                            log.AppendLine($"  'o' separator R{r} → sekcja: {currentLocation}");
                        }
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(col1)) continue;

                    // Etykieta sekcji w wierszu danych — pomiń (tryb A obsłużył powyżej)
                    if (DetectSectionLabel(ws, r, lastCol) != null) continue;

                    if (!IsPileId(col1)) continue;

                    string pileId = col1;
                    string action = ws.Cells[r, colAction].GetValue<string>()?.Trim() ?? "NO ACTION";

                    double util = 0;
                    TryReadUtil(ws, r, lastCol, out util);

                    piles.Add(new PileData
                    {
                        PileId         = pileId,
                        UtilPct        = util,
                        LocationType   = currentLocation,
                        PunchingAction = action
                    });
                }

                log.AppendLine($"Wczytano {piles.Count} pali.");
                foreach (var g in piles.GroupBy(p => p.LocationType))
                    log.AppendLine($"  {g.Key}: {g.Count()} pali — " +
                        $"akcje: {string.Join(", ", g.Select(p => p.PunchingAction).Distinct())}");
            }

            parseLog = log.ToString();
            return piles;
        }

        // ── Pomocnicze ──────────────────────────────────────────────────────────────

        private static int FindActionColumn(ExcelWorksheet ws, int lastRow, int lastCol,
            System.Text.StringBuilder log)
        {
            for (int c = lastCol; c >= 1; c--)
            {
                for (int r = 7; r <= Math.Min(50, lastRow); r++)
                {
                    string v = ws.Cells[r, c].GetValue<string>()?.Trim() ?? "";
                    if (IsActionValue(v)) return c;
                }
            }
            return -1;
        }

        /// <summary>
        /// Sprawdza czy wiersz zawiera etykietę sekcji (INTERNAL/EDGE/CORNER/REENTRANT PILES).
        /// Skanuje kilka pierwszych kolumn (etykieta może być w scalonej komórce).
        /// Zwraca kod lokalizacji lub null.
        /// </summary>
        private static string DetectSectionLabel(ExcelWorksheet ws, int row, int lastCol)
        {
            for (int c = 1; c <= Math.Min(lastCol, 10); c++)
            {
                string v = ws.Cells[row, c].GetValue<string>()?.Trim()?.ToUpperInvariant() ?? "";
                if (string.IsNullOrEmpty(v)) continue;

                // Dopasuj konkretne etykiety sekcji, nie pojedyncze słowa
                if (v.Contains("INTERNAL PILE")) return "INT";
                if (v.Contains("CORNER PILE"))   return "CORNER";
                if (v.Contains("EDGE PILE"))     return "EDGE";
                if (v.Contains("REENTRANT PILE")) return "REENTRANT";
                // Skrócone wersje (cała komórka)
                if (v == "INTERNAL" || v == "INT. PILES" || v == "INT PILES") return "INT";
                if (v == "CORNER"   || v == "CORNER PILES")                   return "CORNER";
                if (v == "EDGE"     || v == "EDGE PILES")                     return "EDGE";
                if (v == "REENTRANT")                                         return "REENTRANT";
            }
            return null;
        }

        /// <summary>
        /// Szuka kolumny z wartością utylizacji (%) w danym wierszu.
        /// Próbuje col 9, a potem skanuje w poszukiwaniu wartości procentowej.
        /// </summary>
        private static void TryReadUtil(ExcelWorksheet ws, int row, int lastCol, out double util)
        {
            util = 0;
            // Najpierw kolumna 9 (standardowa)
            if (TryReadDouble(ws.Cells[row, 9].GetValue<string>(), out util) && util > 0)
                return;

            // Szukaj kolumny z wartością >0 i <200 (realny % utylizacji)
            for (int c = 5; c <= Math.Min(lastCol, 20); c++)
            {
                string v = ws.Cells[row, c].GetValue<string>()?.Trim() ?? "";
                if (TryReadDouble(v, out double d) && d > 0 && d < 200)
                {
                    util = d;
                    return;
                }
            }
        }

        private static void DumpRows(ExcelWorksheet ws, int lastRow, System.Text.StringBuilder log)
        {
            for (int r = 1; r <= Math.Min(10, lastRow); r++)
            {
                var vals = Enumerable.Range(1, 6)
                    .Select(c => ws.Cells[r, c].GetValue<string>() ?? "");
                log.AppendLine($"  R{r}: {string.Join(" | ", vals)}");
            }
        }

        private static bool IsActionValue(string v)
        {
            if (string.IsNullOrEmpty(v)) return false;
            string u = v.ToUpperInvariant();
            return u.StartsWith("ADD H") || u == "NO ACTION";
        }

        private static bool IsPileId(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return false;
            string u = val.Trim().ToUpperInvariant();
            // Odrzuć znane nagłówki i separatory
            if (u.EndsWith("PILES"))     return false;
            if (u.Contains("INTERNAL")) return false;
            if (u.Contains("CORNER"))   return false;
            if (u.Contains("EDGE"))     return false;
            if (u.Contains("REENTRANT")) return false;
            if (u.Contains("REDUC"))    return false;
            if (u.Contains("SECTION"))  return false;
            if (u.Contains("NIB"))      return false;
            if (u == "]" || u == "O" || u == "I" || u == "V" || u == "C") return false;
            // Akceptuj liczby lub krótkie kody alfanumeryczne (pile IDs)
            if (int.TryParse(u, out _)) return true;
            if (u.Length >= 2 && u.Length <= 8 &&
                u.All(ch => char.IsLetterOrDigit(ch) || ch == '-'))
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
