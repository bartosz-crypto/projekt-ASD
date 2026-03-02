using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AsdRcSlab
{
    public static class PunchingParser
    {
        public static List<PileData> Parse(string xlsxPath, out string parseLog)
        {
            var piles = new List<PileData>();
            var log   = new System.Text.StringBuilder();

            using (var pkg = new ExcelPackage(new FileInfo(xlsxPath)))
            {
                // Znajdz wlasciwy arkusz
                ExcelWorksheet ws = FindSheet(pkg);
                if (ws == null)
                {
                    log.AppendLine("Nie znaleziono arkusza z danymi pali.");
                    log.AppendLine("Dostepne arkusze:");
                    foreach (var s in pkg.Workbook.Worksheets)
                        log.AppendLine($"  — {s.Name}");
                    parseLog = log.ToString();
                    return piles;
                }

                log.AppendLine($"Wczytuje arkusz: '{ws.Name}'");

                int lastRow = ws.Dimension?.End.Row ?? 1;
                int lastCol = ws.Dimension?.End.Column ?? 20;

                // Znajdz naglowki
                int headerRow = -1;
                int colId = -1, colUtil = -1, colLoc = -1;

                for (int r = 1; r <= Math.Min(15, lastRow); r++)
                {
                    for (int c = 1; c <= lastCol; c++)
                    {
                        string cell = ws.Cells[r, c].GetValue<string>()?.Trim().ToUpperInvariant() ?? "";
                        if (colId < 0 && (cell == "PILE" || cell == "PILE ID" || cell == "P_NO" ||
                            cell == "PILE NO" || cell == "NO" || cell == "ID"))
                        { headerRow = r; colId = c; }
                        else if (colUtil < 0 && (cell.Contains("UTIL") || cell.Contains("RATIO") ||
                            cell == "U%" || cell == "UTILISATION" || cell == "UTILIZATION"))
                        { colUtil = c; }
                        else if (colLoc < 0 && (cell.Contains("LOCAT") || cell.Contains("TYPE") ||
                            cell == "POSITION" || cell == "PH" || cell == "CONDITION"))
                        { colLoc = c; }
                    }
                    if (headerRow > 0 && colId > 0) break;
                }

                if (headerRow < 0)
                {
                    log.AppendLine("Nie znaleziono naglowkow — szukano: Pile ID, Util%, Location.");
                    log.AppendLine($"Naglowki w wierszu 1:");
                    for (int c = 1; c <= Math.Min(lastCol, 10); c++)
                        log.AppendLine($"  Kol {c}: '{ws.Cells[1, c].GetValue<string>()}'");
                    parseLog = log.ToString();
                    return piles;
                }

                log.AppendLine($"Naglowki: ID=kol{colId}, Util=kol{colUtil}, Loc=kol{colLoc}");

                // Czytaj dane
                for (int r = headerRow + 1; r <= lastRow; r++)
                {
                    string id = ws.Cells[r, colId].GetValue<string>()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(id)) continue;

                    double util = 0;
                    if (colUtil > 0)
                    {
                        string uStr = ws.Cells[r, colUtil].GetValue<string>()?.Trim() ?? "0";
                        uStr = uStr.Replace("%", "").Replace(",", ".");
                        double.TryParse(uStr, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out util);
                    }

                    string loc = "";
                    if (colLoc > 0)
                        loc = NormalizeLocation(ws.Cells[r, colLoc].GetValue<string>()?.Trim() ?? "");

                    piles.Add(new PileData
                    {
                        PileId       = id,
                        UtilPct      = util,
                        LocationType = loc
                    });
                }

                log.AppendLine($"Wczytano {piles.Count} pali.");
            }

            parseLog = log.ToString();
            return piles;
        }

        private static ExcelWorksheet FindSheet(ExcelPackage pkg)
        {
            string[] preferred = { "Summary", "Pile", "PUNCHING", "Results", "Data" };
            foreach (var name in preferred)
                foreach (var ws in pkg.Workbook.Worksheets)
                    if (ws.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                        return ws;

            // Fallback: pierwszy arkusz z wiecej niz 3 wierszami
            return pkg.Workbook.Worksheets
                .FirstOrDefault(w => (w.Dimension?.End.Row ?? 0) > 3)
                ?? pkg.Workbook.Worksheets.FirstOrDefault();
        }

        private static string NormalizeLocation(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "INT";
            string u = raw.ToUpperInvariant();
            if (u.Contains("REENT") || u.Contains("RE-ENT") || u == "RE") return "REENTRANT";
            if (u.Contains("CORN"))  return "CORNER";
            if (u.Contains("EDGE"))  return "EDGE";
            if (u.Contains("INT") || u == "I") return "INT";
            return raw.ToUpperInvariant();
        }
    }
}
