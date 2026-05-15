using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace AsdRcSlab
{
    public static class PunchingParser
    {
        private const string SheetPunchingReport = "Punching Report to Calcs";
        private const int    ColPileId = 1;
        private const int    ColUtil   = 17;
        private const int    ColReinf  = 20;

        private static readonly Regex _plotRx = new Regex(
            @"^PLOT\s+(\d+)(?:\s*-\s*(\d+))?\s*\((\d+)\s*piles?\)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex _sectionRx = new Regex(
            @"^(INTERNAL|EDGE|CORNER|REENTRANT)\s*\((\d+)\)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Łapie formuły cross-sheet: ='Punching EC2'!B60  lub  =Sheet1!B60
        private static readonly Regex CrossSheetRefRx = new Regex(
            @"^=\s*'?([^'!]+)'?\s*!\s*\$?([A-Z]+)\$?(\d+)\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // ── Nowe API (nowy format multi-plot) ───────────────────────────────────

        public static List<PlotInfo> ScanPlots(string xlsxPath, out string log)
        {
            var plots = new List<PlotInfo>();
            var sb    = new StringBuilder();

            using (var pkg = new ExcelPackage(new FileInfo(xlsxPath)))
            {
                var ws = pkg.Workbook.Worksheets[SheetPunchingReport];
                if (ws == null)
                {
                    sb.AppendLine($"Brak arkusza '{SheetPunchingReport}'.");
                    log = sb.ToString();
                    return plots;
                }

                int lastRow = ws.Dimension?.End.Row ?? 1;
                PlotInfo current = null;

                for (int r = 1; r <= lastRow; r++)
                {
                    string c1 = ws.Cells[r, ColPileId].GetValue<string>()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(c1)) continue;

                    var pm = _plotRx.Match(c1);
                    if (pm.Success)
                    {
                        if (current != null) current.EndRow = r - 1;
                        int first     = int.Parse(pm.Groups[1].Value);
                        int last      = pm.Groups[2].Success ? int.Parse(pm.Groups[2].Value) : first;
                        int pileCount = int.Parse(pm.Groups[3].Value);
                        current = new PlotInfo
                        {
                            RawHeader       = c1,
                            FirstPlotNumber = first,
                            LastPlotNumber  = last,
                            PileCount       = pileCount,
                            StartRow        = r
                        };
                        plots.Add(current);
                        continue;
                    }

                    if (current != null)
                    {
                        var sm = _sectionRx.Match(c1);
                        if (sm.Success)
                        {
                            int cnt = int.TryParse(sm.Groups[2].Value, out int n) ? n : 0;
                            switch (sm.Groups[1].Value.ToUpperInvariant())
                            {
                                case "INTERNAL":  current.InternalCount  += cnt; break;
                                case "EDGE":      current.EdgeCount      += cnt; break;
                                case "CORNER":    current.CornerCount    += cnt; break;
                                case "REENTRANT": current.ReentrantCount += cnt; break;
                            }
                        }
                    }
                }

                if (current != null) current.EndRow = lastRow;
                sb.AppendLine($"ScanPlots: znaleziono {plots.Count} plot(ów).");
            }

            log = sb.ToString();
            return plots;
        }

        public static List<PileData> ParsePlot(string xlsxPath, int plotNumber, out string log)
        {
            var piles = new List<PileData>();
            var sb    = new StringBuilder();

            var plots = ScanPlots(xlsxPath, out var scanLog);
            sb.Append(scanLog);

            var plot = plots.FirstOrDefault(p => p.Number == plotNumber);
            if (plot == null)
            {
                sb.AppendLine($"Brak PLOT {plotNumber} w pliku.");
                log = sb.ToString();
                return piles;
            }

            using (var pkg = new ExcelPackage(new FileInfo(xlsxPath)))
            {
                // Krok 1: próba przeliczenia formuł
                TryCalculateWorkbook(pkg, sb);

                var ws = pkg.Workbook.Worksheets[SheetPunchingReport];
                if (ws == null)
                {
                    sb.AppendLine($"Brak arkusza '{SheetPunchingReport}'.");
                    log = sb.ToString();
                    return piles;
                }

                // Krok 2: sprawdź czy cache nie jest pusty (po Calculate + manual resolve)
                var sampleRows = new List<int>();
                for (int r = plot.StartRow + 1; r <= plot.EndRow && sampleRows.Count < 3; r++)
                {
                    string raw = ws.Cells[r, ColPileId].GetValue<string>()?.Trim() ?? "";
                    if (_plotRx.IsMatch(raw) || _sectionRx.IsMatch(raw) ||
                        string.Equals(raw, "Pile", StringComparison.OrdinalIgnoreCase))
                        continue;
                    sampleRows.Add(r);
                }

                if (sampleRows.Count > 0)
                {
                    bool allNull = true;
                    foreach (int sr in sampleRows)
                    {
                        var c1v  = GetCellValue(ws, sr, ColPileId, pkg);
                        var c17v = GetCellValue(ws, sr, ColUtil,   pkg);
                        if (c1v != null || c17v != null) { allNull = false; break; }
                    }
                    if (allNull)
                    {
                        sb.AppendLine("⚠️ Cache formuł pusty nawet po Calculate + manual resolve.");
                        sb.AppendLine("   Otwórz plik w Excelu (Ctrl+S) lub sprawdź czy arkusz źródłowy istnieje.");
                        log = sb.ToString();
                        return piles;
                    }
                }

                // Krok 3: parsuj wiersze danych
                string currentLocation = "INT";
                int parsedCount = 0;

                for (int r = plot.StartRow + 1; r <= plot.EndRow; r++)
                {
                    // Czytaj c1 — użyj GetValue dla nagłówków (tekst literalny), GetCellValue dla danych
                    string c1Literal = ws.Cells[r, ColPileId].GetValue<string>()?.Trim() ?? "";

                    if (_plotRx.IsMatch(c1Literal)) continue;

                    var sm = _sectionRx.Match(c1Literal);
                    if (sm.Success)
                    {
                        currentLocation = NormalizeLocation(sm.Groups[1].Value);
                        continue;
                    }

                    if (string.Equals(c1Literal, "Pile", StringComparison.OrdinalIgnoreCase)) continue;

                    // Wiersz danych — użyj GetCellValue żeby rozwiązać formuły
                    string c1 = Convert.ToString(GetCellValue(ws, r, ColPileId, pkg))?.Trim() ?? "";
                    if (string.IsNullOrEmpty(c1) || c1 == "NaN") continue;

                    double util = 0;
                    try
                    {
                        object v = GetCellValue(ws, r, ColUtil, pkg);
                        if (v is double d)
                            util = d;
                        else if (v != null)
                            double.TryParse(Convert.ToString(v),
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture, out util);
                    }
                    catch { util = 0; }

                    string action = Convert.ToString(GetCellValue(ws, r, ColReinf, pkg))?.Trim() ?? "";

                    var pile = new PileData
                    {
                        PileId         = c1,
                        UtilPct        = util,
                        LocationType   = currentLocation,
                        PunchingAction = action
                    };
                    piles.Add(pile);
                    parsedCount++;

                    // Debug: pierwsze 3 wiersze
                    if (parsedCount <= 3)
                        sb.AppendLine($"  DEBUG row{r}: PileId='{pile.PileId}' Util={pile.UtilPct:F1} " +
                                      $"Reinf='{pile.PunchingAction}' Location='{pile.LocationType}'");
                }

                // Info o nietypowych wartościach Reinf
                var unknownReinf = piles
                    .Select(p => p.PunchingAction)
                    .Where(a => !string.IsNullOrEmpty(a)
                                && !a.StartsWith("ADD H", StringComparison.OrdinalIgnoreCase)
                                && !a.Equals("NO ACTION", StringComparison.OrdinalIgnoreCase))
                    .Distinct().ToList();
                if (unknownReinf.Any())
                    sb.AppendLine($"  INFO: nietypowe wartości Reinf: {string.Join(", ", unknownReinf.Select(s => $"'{s}'"))}");

                int intCnt  = piles.Count(p => p.LocationType == "INT");
                int edgeCnt = piles.Count(p => p.LocationType == "EDGE");
                int cornCnt = piles.Count(p => p.LocationType == "CORNER");
                int reeCnt  = piles.Count(p => p.LocationType == "REENTRANT");
                sb.AppendLine($"ParsePlot {plotNumber}: {piles.Count} pali (INT:{intCnt} EDGE:{edgeCnt} CORNER:{cornCnt} REENTRANT:{reeCnt})");
            }

            log = sb.ToString();
            return piles;
        }

        // ── Stare API (single-plot, zachowane dla zgodności) ────────────────────

        public static List<string> GetSheetNames(string xlsxPath)
        {
            using (var pkg = new ExcelPackage(new FileInfo(xlsxPath)))
                return pkg.Workbook.Worksheets.Select(w => w.Name).ToList();
        }

        [System.Obsolete("Use ScanPlots + ParsePlot for new multi-plot Excel format. Kept for legacy single-plot files.")]
        public static List<PileData> Parse(string xlsxPath, string sheetName, out string parseLog)
        {
            var piles = new List<PileData>();
            var log   = new StringBuilder();

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

                int colAction = FindActionColumn(ws, lastRow, lastCol, log);
                if (colAction < 0)
                {
                    log.AppendLine("Nie znaleziono kolumny ACTION (ADD H.../NO ACTION).");
                    DumpRows(ws, lastRow, log);
                    parseLog = log.ToString();
                    return piles;
                }
                log.AppendLine($"Kolumna ACTION: {colAction}");

                var sectionBoundaries = new List<(int Row, string Location)>();
                for (int r = 1; r <= lastRow; r++)
                {
                    string loc = DetectSectionLabel(ws, r, lastCol);
                    if (loc != null)
                        sectionBoundaries.Add((r, loc));
                }

                sectionBoundaries = sectionBoundaries.OrderBy(s => s.Row).ToList();
                log.AppendLine($"Wykryte sekcje: {string.Join(", ", sectionBoundaries.Select(s => $"R{s.Row}={s.Location}"))}");

                bool dataLabels = sectionBoundaries.Any(s => s.Row > 6);
                var headerOrder = sectionBoundaries.Where(s => s.Row <= 6).Select(s => s.Location).ToList();
                if (headerOrder.Count == 0 && !dataLabels) headerOrder.Add("INT");

                string currentLocation = sectionBoundaries.Count > 0 ? sectionBoundaries[0].Location : "INT";
                int secIdx = 0, headerSecIdx = 0;

                for (int r = 7; r <= lastRow; r++)
                {
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

                    if (col1 == "o" || col1 == "O")
                    {
                        if (!dataLabels && headerOrder.Count > 0)
                        {
                            headerSecIdx = Math.Min(headerSecIdx + 1, headerOrder.Count - 1);
                            currentLocation = headerOrder[headerSecIdx];
                            log.AppendLine($"  'o' separator R{r} → sekcja: {currentLocation}");
                        }
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(col1)) continue;
                    if (DetectSectionLabel(ws, r, lastCol) != null) continue;
                    if (!IsPileId(col1)) continue;

                    string action = ws.Cells[r, colAction].GetValue<string>()?.Trim() ?? "NO ACTION";
                    double util = 0;
                    TryReadUtil(ws, r, lastCol, out util);

                    piles.Add(new PileData
                    {
                        PileId         = col1,
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

        // ── Resolver formuł ─────────────────────────────────────────────────────

        // Próbuje przeliczyć formuły programowo. Best-effort — niektóre funkcje
        // EPPlus 4.5 nie obsługuje. Po wywołaniu cache może (ale nie musi) być
        // wypełniony. Wywołujący nie polega na sukcesie.
        private static void TryCalculateWorkbook(ExcelPackage pkg, StringBuilder log)
        {
            try
            {
                pkg.Workbook.Calculate();
                log.AppendLine("Calculate: OK");
            }
            catch (Exception ex)
            {
                log.AppendLine($"Calculate: FAIL ({ex.GetType().Name}: {ex.Message})");
            }
        }

        // Czyta wartość z komórki. Jeśli Value jest null a Formula to prosty
        // cross-sheet reference, resolvuje go ręcznie przez podstawienie.
        // Zwraca null dla pustych komórek i double.NaN.
        private static object GetCellValue(ExcelWorksheet ws, int row, int col, ExcelPackage pkg)
        {
            var cell = ws.Cells[row, col];

            if (cell.Value != null)
            {
                // Guard przed double.NaN — EPPlus może to zwrócić dla pustej formuły
                if (cell.Value is double d && double.IsNaN(d)) return null;
                return cell.Value;
            }

            if (string.IsNullOrEmpty(cell.Formula)) return null;

            // EPPlus zwraca formułę BEZ wiodącego '=' — prependujemy żeby regex łapał
            var m = CrossSheetRefRx.Match("=" + cell.Formula);
            if (!m.Success) return null;

            string targetSheetName = m.Groups[1].Value;
            string targetAddr      = m.Groups[2].Value + m.Groups[3].Value;

            var targetSheet = pkg.Workbook.Worksheets[targetSheetName];
            if (targetSheet == null) return null;

            var resolved = targetSheet.Cells[targetAddr].Value;
            if (resolved is double rd && double.IsNaN(rd)) return null;
            return resolved;
        }

        // ── Pomocnicze ──────────────────────────────────────────────────────────

        private static string NormalizeLocation(string raw)
        {
            string u = (raw ?? "").ToUpperInvariant().Trim();
            if (u.StartsWith("INTERNAL"))  return "INT";
            if (u.StartsWith("EDGE"))      return "EDGE";
            if (u.StartsWith("CORNER"))    return "CORNER";
            if (u.StartsWith("REENTRANT")) return "REENTRANT";
            return u;
        }

        private static int FindActionColumn(ExcelWorksheet ws, int lastRow, int lastCol,
            StringBuilder log)
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

        private static string DetectSectionLabel(ExcelWorksheet ws, int row, int lastCol)
        {
            for (int c = 1; c <= Math.Min(lastCol, 10); c++)
            {
                string v = ws.Cells[row, c].GetValue<string>()?.Trim()?.ToUpperInvariant() ?? "";
                if (string.IsNullOrEmpty(v)) continue;

                if (v.Contains("INTERNAL PILE"))  return "INT";
                if (v.Contains("CORNER PILE"))    return "CORNER";
                if (v.Contains("EDGE PILE"))      return "EDGE";
                if (v.Contains("REENTRANT PILE")) return "REENTRANT";
                if (v == "INTERNAL" || v == "INT. PILES" || v == "INT PILES") return "INT";
                if (v == "CORNER"   || v == "CORNER PILES")                   return "CORNER";
                if (v == "EDGE"     || v == "EDGE PILES")                     return "EDGE";
                if (v == "REENTRANT")                                         return "REENTRANT";
            }
            return null;
        }

        private static void TryReadUtil(ExcelWorksheet ws, int row, int lastCol, out double util)
        {
            util = 0;
            if (TryReadDouble(ws.Cells[row, 9].GetValue<string>(), out util) && util > 0)
                return;

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

        private static void DumpRows(ExcelWorksheet ws, int lastRow, StringBuilder log)
        {
            for (int r = 1; r <= Math.Min(10, lastRow); r++)
            {
                var vals = Enumerable.Range(1, 6).Select(c => ws.Cells[r, c].GetValue<string>() ?? "");
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
            if (u.EndsWith("PILES"))      return false;
            if (u.Contains("INTERNAL"))  return false;
            if (u.Contains("CORNER"))    return false;
            if (u.Contains("EDGE"))      return false;
            if (u.Contains("REENTRANT")) return false;
            if (u.Contains("REDUC"))     return false;
            if (u.Contains("SECTION"))   return false;
            if (u.Contains("NIB"))       return false;
            if (u == "]" || u == "O" || u == "I" || u == "V" || u == "C") return false;
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
