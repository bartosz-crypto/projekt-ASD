using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            @"^PLOT\s+(\d+)\s*\((\d+)\s*piles?\)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex _sectionRx = new Regex(
            @"^(INTERNAL|EDGE|CORNER|REENTRANT)\s*\((\d+)\)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // ── Nowe API (nowy format multi-plot) ───────────────────────────────────

        public static List<PlotInfo> ScanPlots(string xlsxPath, out string log)
        {
            var plots = new List<PlotInfo>();
            var sb    = new System.Text.StringBuilder();

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
                        current = new PlotInfo
                        {
                            Number    = int.Parse(pm.Groups[1].Value),
                            PileCount = int.Parse(pm.Groups[2].Value),
                            StartRow  = r
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
            var sb    = new System.Text.StringBuilder();

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
                var ws = pkg.Workbook.Worksheets[SheetPunchingReport];
                if (ws == null)
                {
                    sb.AppendLine($"Brak arkusza '{SheetPunchingReport}'.");
                    log = sb.ToString();
                    return piles;
                }

                // Sprawdź czy cache formuł nie jest pusty
                int checkedRows = 0, nullRows = 0;
                for (int r = plot.StartRow + 1; r <= plot.EndRow && checkedRows < 5; r++)
                {
                    string c1chk = ws.Cells[r, ColPileId].GetValue<string>()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(c1chk)) continue;
                    if (_plotRx.IsMatch(c1chk) || _sectionRx.IsMatch(c1chk) ||
                        string.Equals(c1chk, "Pile", StringComparison.OrdinalIgnoreCase)) continue;

                    checkedRows++;
                    string c17chk = ws.Cells[r, ColUtil].GetValue<string>()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(c17chk)) nullRows++;
                }
                if (checkedRows > 0 && nullRows == checkedRows)
                {
                    sb.AppendLine("Cache formuł pusty — plik nie był przeliczony w Excelu.");
                    log = sb.ToString();
                    return piles;
                }

                string currentLocation = "INT";

                for (int r = plot.StartRow + 1; r <= plot.EndRow; r++)
                {
                    string c1 = ws.Cells[r, ColPileId].GetValue<string>()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(c1)) continue;

                    if (_plotRx.IsMatch(c1)) continue;

                    var sm = _sectionRx.Match(c1);
                    if (sm.Success)
                    {
                        currentLocation = NormalizeLocation(sm.Groups[1].Value);
                        continue;
                    }

                    if (string.Equals(c1, "Pile", StringComparison.OrdinalIgnoreCase)) continue;

                    // Wiersz danych — odczyt Util z c17
                    string rawUtil = ws.Cells[r, ColUtil].GetValue<string>()?.Trim() ?? "";
                    double util = 0;
                    if (TryReadDouble(rawUtil, out double rawVal) && rawVal > 0)
                        util = rawVal < 5.0 ? rawVal * 100.0 : rawVal;  // ułamek lub procent

                    string action = ws.Cells[r, ColReinf].GetValue<string>()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(action)) action = "NO ACTION";

                    piles.Add(new PileData
                    {
                        PileId         = c1,
                        UtilPct        = util,
                        LocationType   = currentLocation,
                        PunchingAction = action
                    });
                }

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

        private static void DumpRows(ExcelWorksheet ws, int lastRow, System.Text.StringBuilder log)
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
