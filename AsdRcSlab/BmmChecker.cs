using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AsdRcSlab
{
    public class BmmRuleResult
    {
        public string RuleId  { get; set; }
        public string Status  { get; set; }  // OK / WARN / FAIL
        public string Details { get; set; }
    }

    public class BmmResult
    {
        public BmmRuleResult R87 { get; set; }  // Luki w numeracji
        public BmmRuleResult R95 { get; set; }  // Multipliers > 1
        public BmmRuleResult R81 { get; set; }  // Min dia BOTTOM
        public BmmRuleResult R83 { get; set; }  // Min dia TOP
        public BmmRuleResult R92 { get; set; }  // Kolumna Type wypelniona
    }

    public static class BmmChecker
    {
        public static BmmResult CheckAll(string xlsxPath)
        {
            using (var package = new ExcelPackage(new FileInfo(xlsxPath)))
            {
                var wsBottom = FindSheet(package, "BOTTOM", "Bot", "B1");
                var wsTop    = FindSheet(package, "TOP",    "Top", "T1");

                var bottomRows = wsBottom != null ? ReadBarsFromSheet(wsBottom) : new List<BarRow>();
                var topRows    = wsTop    != null ? ReadBarsFromSheet(wsTop)    : new List<BarRow>();
                var allRows    = bottomRows.Concat(topRows).ToList();

                return new BmmResult
                {
                    R87 = CheckGaps(allRows),
                    R95 = CheckMultipliers(allRows),
                    R81 = CheckMinDia(bottomRows, minDia: 10, ruleId: "R81", label: "BOTTOM"),
                    R83 = CheckMinDia(topRows,    minDia: 12, ruleId: "R83", label: "TOP"),
                    R92 = CheckTypeFilled(allRows)
                };
            }
        }

        // ── Pomocnicze ───────────────────────────────────────────────────────

        private static ExcelWorksheet FindSheet(ExcelPackage pkg, params string[] candidates)
        {
            foreach (var ws in pkg.Workbook.Worksheets)
                foreach (var name in candidates)
                    if (ws.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                        return ws;
            return null;
        }

        private static List<BarRow> ReadBarsFromSheet(ExcelWorksheet ws)
        {
            var rows = new List<BarRow>();
            int lastRow = ws.Dimension?.End.Row ?? 1;
            int lastCol = ws.Dimension?.End.Column ?? 20;

            // Znajdz wiersz naglowkowy
            int headerRow = -1, colMark = -1, colType = -1, colNoMbrs = -1;

            for (int r = 1; r <= Math.Min(10, lastRow); r++)
            {
                for (int c = 1; c <= lastCol; c++)
                {
                    string cell = ws.Cells[r, c].GetValue<string>()?.Trim() ?? "";
                    if (cell.Equals("Mark", StringComparison.OrdinalIgnoreCase) ||
                        cell.Equals("Bar Mark", StringComparison.OrdinalIgnoreCase))
                    { headerRow = r; colMark = c; }
                    else if (cell.Equals("Type", StringComparison.OrdinalIgnoreCase) ||
                             cell.Equals("Bar Type", StringComparison.OrdinalIgnoreCase) ||
                             cell.Equals("Size", StringComparison.OrdinalIgnoreCase))
                    { colType = c; }
                    else if (cell.StartsWith("No", StringComparison.OrdinalIgnoreCase) &&
                             cell.IndexOf("Mbr", StringComparison.OrdinalIgnoreCase) >= 0)
                    { colNoMbrs = c; }
                }
                if (headerRow > 0) break;
            }

            if (headerRow < 0 || colMark < 0) return rows;

            for (int r = headerRow + 1; r <= lastRow; r++)
            {
                string markStr = ws.Cells[r, colMark].GetValue<string>()?.Trim() ?? "";
                if (string.IsNullOrEmpty(markStr)) continue;

                var bar = new BarRow { MarkRaw = markStr };

                if (int.TryParse(markStr, out int markInt))
                    bar.Mark = markInt;

                if (colType > 0)
                    bar.TypeRaw = ws.Cells[r, colType].GetValue<string>()?.Trim() ?? "";

                if (colNoMbrs > 0)
                {
                    string nmStr = ws.Cells[r, colNoMbrs].GetValue<string>()?.Trim() ?? "1";
                    if (int.TryParse(nmStr, out int nm)) bar.NoMbrs = nm;
                }

                rows.Add(bar);
            }

            return rows;
        }

        // ── Reguly ──────────────────────────────────────────────────────────

        private static BmmRuleResult CheckGaps(List<BarRow> rows)
        {
            var marks = rows
                .Where(r => r.Mark.HasValue)
                .Select(r => r.Mark.Value)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            if (!marks.Any())
                return new BmmRuleResult { RuleId = "R87", Status = "WARN",
                    Details = "Brak numerycznych oznaczeń prętów — sprawdź format kolumny Mark." };

            var gaps = new List<string>();
            for (int i = marks[0]; i <= marks[marks.Count - 1]; i++)
                if (!marks.Contains(i))
                    gaps.Add(i.ToString("D2"));

            if (!gaps.Any())
                return new BmmRuleResult { RuleId = "R87", Status = "OK",
                    Details = $"Brak luk. Pręty {marks.First():D2}–{marks.Last():D2} ({marks.Count} szt.)" };

            return new BmmRuleResult { RuleId = "R87", Status = "FAIL",
                Details = $"Luki w numeracji — brakuje nr: {string.Join(", ", gaps)}" };
        }

        private static BmmRuleResult CheckMultipliers(List<BarRow> rows)
        {
            var multi = rows.Where(r => r.NoMbrs > 1).ToList();
            if (!multi.Any())
                return new BmmRuleResult { RuleId = "R95", Status = "OK",
                    Details = "Wszystkie No.Mbrs = 1." };

            var list = string.Join(", ", multi.Select(r => $"Mark {r.MarkRaw} (×{r.NoMbrs})"));
            return new BmmRuleResult { RuleId = "R95", Status = "WARN",
                Details = $"No.Mbrs > 1: {list}" };
        }

        private static BmmRuleResult CheckMinDia(List<BarRow> rows, int minDia,
            string ruleId, string label)
        {
            var fails = rows
                .Where(r => { int d = ParseDia(r.TypeRaw); return d > 0 && d < minDia; })
                .Select(r => $"Mark {r.MarkRaw} (H{ParseDia(r.TypeRaw)})")
                .ToList();

            if (!fails.Any())
                return new BmmRuleResult { RuleId = ruleId, Status = "OK",
                    Details = $"{label}: wszystkie pręty ≥ H{minDia}." };

            return new BmmRuleResult { RuleId = ruleId, Status = "FAIL",
                Details = $"{label} min Ø H{minDia} — niezgodne: {string.Join(", ", fails)}" };
        }

        private static BmmRuleResult CheckTypeFilled(List<BarRow> rows)
        {
            int noType = rows.Count(r => string.IsNullOrEmpty(r.TypeRaw));
            if (noType > rows.Count / 2)
                return new BmmRuleResult { RuleId = "R92", Status = "WARN",
                    Details = $"Brak danych Type dla {noType} prętów — nie można wyliczyć masy." };

            return new BmmRuleResult { RuleId = "R92", Status = "OK",
                Details = "Kolumna Type wypełniona — waga weryfikowalna." };
        }

        private static int ParseDia(string type)
        {
            if (string.IsNullOrEmpty(type)) return 0;
            string digits = new string(type.Where(char.IsDigit).ToArray());
            return int.TryParse(digits, out int d) ? d : 0;
        }
    }

    public class BarRow
    {
        public string MarkRaw { get; set; } = "";
        public int?   Mark    { get; set; }
        public string TypeRaw { get; set; } = "";
        public int    NoMbrs  { get; set; } = 1;
    }
}
