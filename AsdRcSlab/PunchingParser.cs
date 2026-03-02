using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AsdRcSlab
{
    public static class PunchingParser
    {
        // Zwraca liste nazw arkuszy z pliku
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

                int lastRow = ws.Dimension?.End.Row ?? 1;
                int lastCol = ws.Dimension?.End.Column ?? 50;

                log.AppendLine($"Arkusz: '{ws.Name}', wiersze: {lastRow}, kolumny: {lastCol}");

                // Znajdz wiersz naglowkowy — szukaj "Pile id" lub "Pile ID"
                int headerRow = -1;
                int colId = -1, colShear = -1, colUtil = -1;

                for (int r = 1; r <= Math.Min(20, lastRow); r++)
                {
                    for (int c = 1; c <= Math.Min(lastCol, 60); c++)
                    {
                        string val = ws.Cells[r, c].GetValue<string>()?.Trim() ?? "";

                        if (colId < 0 &&
                            (val.Equals("Pile id", StringComparison.OrdinalIgnoreCase) ||
                             val.Equals("Pile ID", StringComparison.OrdinalIgnoreCase) ||
                             val.Equals("Pile No", StringComparison.OrdinalIgnoreCase)))
                        { headerRow = r; colId = c; }

                        if (colShear < 0 &&
                            val.IndexOf("SHEAR CONDITION", StringComparison.OrdinalIgnoreCase) >= 0)
                        { colShear = c; }

                        if (colUtil < 0 && val.Trim() == "%")
                        { colUtil = c; }
                    }
                    if (headerRow > 0 && colShear > 0 && colUtil > 0) break;
                }

                if (headerRow < 0)
                {
                    log.AppendLine("Nie znaleziono wiersza nagłówkowego (szukano 'Pile id').");
                    log.AppendLine("Pierwsze 5 wierszy kol 1-5:");
                    for (int r = 1; r <= Math.Min(5, lastRow); r++)
                    {
                        var vals = Enumerable.Range(1, 5)
                            .Select(c => ws.Cells[r, c].GetValue<string>() ?? "");
                        log.AppendLine($"  R{r}: {string.Join(" | ", vals)}");
                    }
                    parseLog = log.ToString();
                    return piles;
                }

                log.AppendLine($"Nagłówek: wiersz {headerRow}, Pile id=kol{colId}, ShearCond=kol{colShear}, %=kol{colUtil}");

                // Czytaj wiersze z danymi
                for (int r = headerRow + 1; r <= lastRow; r++)
                {
                    string id = ws.Cells[r, colId].GetValue<string>()?.Trim() ?? "";

                    // Pomijaj puste, sekcje ("EDGE PILES", "INTERNAL PILES" itp.)
                    if (string.IsNullOrWhiteSpace(id)) continue;
                    if (id.EndsWith("PILES", StringComparison.OrdinalIgnoreCase)) continue;
                    if (id.Equals("Pile id", StringComparison.OrdinalIgnoreCase)) continue;

                    // Util %
                    double util = 0;
                    if (colUtil > 0)
                    {
                        string uStr = ws.Cells[r, colUtil].GetValue<string>()?.Trim()
                                      ?? ws.Cells[r, colUtil].GetValue<double>().ToString();
                        uStr = uStr.Replace("%", "").Replace(",", ".").Trim();
                        double.TryParse(uStr, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out util);
                    }

                    // Location — z kolumny SHEAR CONDITION
                    string loc = "";
                    if (colShear > 0)
                    {
                        string raw = ws.Cells[r, colShear].GetValue<string>()?.Trim() ?? "";
                        loc = NormalizeLocation(raw);
                    }

                    if (string.IsNullOrEmpty(loc)) loc = "INT"; // fallback

                    piles.Add(new PileData
                    {
                        PileId       = id,
                        UtilPct      = util,
                        LocationType = loc
                    });
                }

                log.AppendLine($"Wczytano {piles.Count} pali.");
                if (piles.Count > 0)
                {
                    int edge   = piles.Count(p => p.LocationType == "EDGE");
                    int corner = piles.Count(p => p.LocationType == "CORNER");
                    int intern = piles.Count(p => p.LocationType == "INT");
                    int re     = piles.Count(p => p.LocationType == "REENTRANT");
                    log.AppendLine($"  INT={intern}, EDGE={edge}, CORNER={corner}, REENTRANT={re}");
                }
            }

            parseLog = log.ToString();
            return piles;
        }

        private static string NormalizeLocation(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "";
            string u = raw.Trim().ToUpperInvariant();

            // Format z PLOT.xlsx: "E", "C", "I"
            if (u == "E")  return "EDGE";
            if (u == "C")  return "CORNER";
            if (u == "I")  return "INT";
            if (u == "RE") return "REENTRANT";

            // Pelne slowa
            if (u.Contains("REENT") || u.Contains("RE-ENT")) return "REENTRANT";
            if (u.Contains("CORN"))  return "CORNER";
            if (u.Contains("EDGE"))  return "EDGE";
            if (u.Contains("INT"))   return "INT";

            return u;
        }
    }
}
