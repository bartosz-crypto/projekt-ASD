using OfficeOpenXml;
using OfficeOpenXml.Style;
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

        // p162: granica PŁYTY = komórka kol. A z fill SOLID o kolorze NIEBIESKIM 0D47A1
        // (zweryfikowane na 3 plikach: kolor występuje WYŁĄCZNIE na nagłówkach płyt).
        private const string SlabHeaderFillRgb = "0D47A1";

        // Sufiks "(N piles)" w nagłówku płyty — do wyciągnięcia PileCount i oczyszczenia Label.
        private static readonly Regex _pilesRx = new Regex(
            @"\(\s*(\d+)\s*piles?\s*\)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Nagłówek plotu — separatoro-ODPORNY. "spec" zaczyna się cyfrą i może zawierać
        // cyfry, myślniki (zakresy) oraz separatory ( , ; : / & spacja kropka ).
        // Łapie: "PLOT 1", "PLOT 4-5", "PLOT 10, 32", "PLOT 14-15, 28-29",
        //        "PLOT 18; 25; 27", "PLOT 10 : 32", "PLOT 10 / 32" — wszystkie z "(k piles)".
        // Właściwe parsowanie liczb robi ParsePlotNumbers().
        private static readonly Regex _plotRx = new Regex(
            @"^PLOT\s+(?<spec>\d[\d\s,;:/&.\-]*?)\s*\((?<count>\d+)\s*piles?\)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Token = pojedyncza liczba lub zakres "a-b". Stosowany na "spec".
        private static readonly Regex _plotTokenRx = new Regex(
            @"\d+(?:\s*-\s*\d+)?", RegexOptions.Compiled);

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

                // p162: granica płyty = niebieski (0D47A1) SOLID fill w kol. A — nie słowo "PLOT".
                ScanByColor(ws, lastRow, plots, sb);

                // FALLBACK 1: raport bez koloru → spróbuj legacy _plotRx (nie ciche mis-parsowanie).
                if (plots.Count == 0)
                {
                    sb.AppendLine("⚠️ p162: nie znaleziono ŻADNEGO niebieskiego nagłówka płyty " +
                                  "(fill SOLID 0D47A1 w kol. A). Raport bez koloru? Próbuję legacy _plotRx.");
                    ScanByPlotRx(ws, lastRow, plots, sb);
                }

                // FALLBACK 2: brak jakichkolwiek nagłówków → cały arkusz jako 1 plot.
                if (plots.Count == 0)
                {
                    int intCnt = 0, edgeCnt = 0, cornerCnt = 0, reentrantCnt = 0;
                    for (int r = 1; r <= lastRow; r++)
                    {
                        string c1f = ws.Cells[r, ColPileId].GetValue<string>()?.Trim() ?? "";
                        if (string.IsNullOrEmpty(c1f)) continue;
                        var sm = _sectionRx.Match(c1f);
                        if (!sm.Success) continue;
                        int cnt = int.TryParse(sm.Groups[2].Value, out int n) ? n : 0;
                        switch (sm.Groups[1].Value.ToUpperInvariant())
                        {
                            case "INTERNAL":  intCnt       += cnt; break;
                            case "EDGE":      edgeCnt      += cnt; break;
                            case "CORNER":    cornerCnt    += cnt; break;
                            case "REENTRANT": reentrantCnt += cnt; break;
                        }
                    }
                    int totalPiles = intCnt + edgeCnt + cornerCnt + reentrantCnt;
                    if (totalPiles > 0)
                    {
                        plots.Add(new PlotInfo
                        {
                            RawHeader       = "SYNTHETIC (no slab header)",
                            Label           = "PLOT 1",
                            PlotNumbers     = new List<int> { 1 },
                            PileCount       = totalPiles,
                            StartRow        = 1,
                            EndRow          = lastRow,
                            InternalCount   = intCnt,
                            EdgeCount       = edgeCnt,
                            CornerCount     = cornerCnt,
                            ReentrantCount  = reentrantCnt
                        });
                        sb.AppendLine($"FALLBACK: brak nagłówków płyt — syntetyczny PLOT 1: {totalPiles} pali " +
                                      $"(INT:{intCnt} EDGE:{edgeCnt} CORNER:{cornerCnt} REENTRANT:{reentrantCnt}).");
                    }
                }

                sb.AppendLine($"ScanPlots: znaleziono {plots.Count} plot(ów).");
            }

            log = sb.ToString();
            return plots;
        }

        // p162: skan po KOLORZE — granica płyty = IsSlabHeaderRow (niebieski 0D47A1).
        // Label = tekst kol. A bez sufiksu "(N piles)"; PileCount z sufiksu; PlotNumbers
        // best-effort z Label (DOPUSZCZALNE PUSTE dla nazw typu "CARE HOME"). Wiersze przed
        // pierwszym nagłówkiem = preambuła (pomijane). Sekcje teal akumulują się w bieżącej płycie.
        private static void ScanByColor(ExcelWorksheet ws, int lastRow, List<PlotInfo> plots, StringBuilder sb)
        {
            PlotInfo current = null;

            for (int r = 1; r <= lastRow; r++)
            {
                if (IsSlabHeaderRow(ws, r))
                {
                    if (current != null) current.EndRow = r - 1;
                    string c1    = ws.Cells[r, ColPileId].GetValue<string>()?.Trim() ?? "";
                    string label = StripPilesSuffix(c1);
                    current = new PlotInfo
                    {
                        RawHeader   = c1,
                        Label       = label,
                        PlotNumbers = ParsePlotNumbers(label),   // best-effort; może być pusta
                        PileCount   = ExtractPileCount(c1),
                        StartRow    = r
                    };
                    plots.Add(current);
                    continue;
                }

                if (current == null) continue;   // preambuła przed pierwszym nagłówkiem

                string txt = ws.Cells[r, ColPileId].GetValue<string>()?.Trim() ?? "";
                if (string.IsNullOrEmpty(txt)) continue;

                var sm = _sectionRx.Match(txt);
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

            if (current != null) current.EndRow = lastRow;
        }

        // Legacy: granica po _plotRx (słowo "PLOT N (k piles)"). Fallback gdy brak koloru.
        private static void ScanByPlotRx(ExcelWorksheet ws, int lastRow, List<PlotInfo> plots, StringBuilder sb)
        {
            PlotInfo current = null;

            for (int r = 1; r <= lastRow; r++)
            {
                string c1 = ws.Cells[r, ColPileId].GetValue<string>()?.Trim() ?? "";
                if (string.IsNullOrEmpty(c1)) continue;

                var pm = _plotRx.Match(c1);
                if (pm.Success)
                {
                    if (current != null) current.EndRow = r - 1;
                    string spec   = pm.Groups["spec"].Value;
                    current = new PlotInfo
                    {
                        RawHeader   = c1,
                        Label       = "PLOT " + spec.Trim(),
                        PlotNumbers = ParsePlotNumbers(spec),
                        PileCount   = int.Parse(pm.Groups["count"].Value),
                        StartRow    = r
                    };
                    plots.Add(current);
                    continue;
                }

                if (current == null) continue;

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

            if (current != null) current.EndRow = lastRow;
        }

        // p162: granica płyty = kol. A ma fill SOLID i RGB kończy się na 0D47A1.
        // Kolor jest dyskryminatorem (sekcje teal też są scalone+bold). Null/"" Rgb => nie nagłówek.
        private static bool IsSlabHeaderRow(ExcelWorksheet ws, int r)
        {
            var fill = ws.Cells[r, ColPileId].Style.Fill;
            if (fill.PatternType != ExcelFillStyle.Solid) return false;
            var rgb = fill.BackgroundColor?.Rgb;            // ARGB, np. "FF0D47A1"/"000D47A1"
            return !string.IsNullOrEmpty(rgb) && rgb.Length >= 6 &&
                   rgb.Substring(rgb.Length - 6).Equals(SlabHeaderFillRgb, StringComparison.OrdinalIgnoreCase);
        }

        private static int ExtractPileCount(string s)
        {
            var m = _pilesRx.Match(s ?? "");
            return m.Success && int.TryParse(m.Groups[1].Value, out int n) ? n : 0;
        }

        private static string StripPilesSuffix(string s)
        {
            if (string.IsNullOrEmpty(s)) return s ?? "";
            return _pilesRx.Replace(s, "").Trim();
        }

        // p162: parsowanie wybranego plotu BEZ matchu po Number — czyta wprost po StartRow/EndRow.
        // Eliminuje kolizję dla płyt bez numerów (np. "CARE HOME PART 1/2" i "...2/2" → oba Number==0).
        public static List<PileData> ParsePlot(string xlsxPath, PlotInfo plot, out string log)
        {
            var sb = new StringBuilder();
            if (plot == null)
            {
                sb.AppendLine("ParsePlot: plot == null.");
                log = sb.ToString();
                return new List<PileData>();
            }
            var piles = ParsePlotRows(xlsxPath, plot, sb);
            log = sb.ToString();
            return piles;
        }

        // Backward-compat: lookup po Number (FirstPlotNumber). Niebezpieczne dla płyt bez
        // numerów — preferuj przeciążenie ParsePlot(xlsxPath, PlotInfo, ...).
        public static List<PileData> ParsePlot(string xlsxPath, int plotNumber, out string log)
        {
            var sb    = new StringBuilder();
            var plots = ScanPlots(xlsxPath, out var scanLog);
            sb.Append(scanLog);

            var plot = plots.FirstOrDefault(p => p.Number == plotNumber);
            if (plot == null)
            {
                sb.AppendLine($"Brak PLOT {plotNumber} w pliku.");
                log = sb.ToString();
                return new List<PileData>();
            }

            var piles = ParsePlotRows(xlsxPath, plot, sb);
            log = sb.ToString();
            return piles;
        }

        // Wspólne czytanie wierszy danych po zakresie StartRow+1..EndRow (logika bez zmian).
        private static List<PileData> ParsePlotRows(string xlsxPath, PlotInfo plot, StringBuilder sb)
        {
            var piles = new List<PileData>();

            using (var pkg = new ExcelPackage(new FileInfo(xlsxPath)))
            {
                // Krok 1: próba przeliczenia formuł
                TryCalculateWorkbook(pkg, sb);

                var ws = pkg.Workbook.Worksheets[SheetPunchingReport];
                if (ws == null)
                {
                    sb.AppendLine($"Brak arkusza '{SheetPunchingReport}'.");
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
                sb.AppendLine($"ParsePlot [{plot.Label}]: {piles.Count} pali (INT:{intCnt} EDGE:{edgeCnt} CORNER:{cornCnt} REENTRANT:{reeCnt})");
            }

            return piles;
        }

        // ── Stare API (single-plot, zachowane dla zgodności) ────────────────────

        public static List<string> GetSheetNames(string xlsxPath)
        {
            using (var pkg = new ExcelPackage(new FileInfo(xlsxPath)))
                return pkg.Workbook.Worksheets.Select(w => w.Name).ToList();
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

        // Wyciąga numery plotu z "spec" (separatoro-odpornie). Ignoruje to, co stoi
        // między tokenami. Token = liczba lub zakres "a-b" (rozwijany w pełni).
        // Zwraca uporządkowaną, distinct listę. Przykłady:
        //   "4-5" -> [4,5]; "10, 32" -> [10,32]; "14-15, 28-29" -> [14,15,28,29];
        //   "18 ; 25 ; 27" -> [18,25,27]; "10 / 32" -> [10,32].
        private static List<int> ParsePlotNumbers(string spec)
        {
            var set = new SortedSet<int>();
            if (!string.IsNullOrEmpty(spec))
            {
                foreach (Match t in _plotTokenRx.Matches(spec))
                {
                    string tok = t.Value;
                    int dash = tok.IndexOf('-');
                    if (dash >= 0)
                    {
                        int a = int.Parse(tok.Substring(0, dash).Trim());
                        int b = int.Parse(tok.Substring(dash + 1).Trim());
                        if (a > b) { int tmp = a; a = b; b = tmp; }
                        for (int n = a; n <= b; n++) set.Add(n);
                    }
                    else
                    {
                        set.Add(int.Parse(tok.Trim()));
                    }
                }
            }
            return set.ToList();
        }

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
