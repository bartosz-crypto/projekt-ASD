using System;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace AsdRcSlab
{
    /// <summary>
    /// Czyta plik BBS (.xls lub .xlsx) przez NPOI i wykrywa layout:
    /// gdzie są sekcje BOTTOM LAYER / TOP LAYER w każdym arkuszu.
    /// </summary>
    public static class BbsXlsReader
    {
        private const int ColumnA = 0;        // 0-indexed (NPOI)

        public static BbsXlsLayoutInfo DetectLayout(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path empty.", "path");
            if (!File.Exists(path))
                throw new FileNotFoundException("BBS not found.", path);

            var info = new BbsXlsLayoutInfo { FilePath = path };

            using (var fs = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IWorkbook wb = OpenWorkbook(fs, path);

                for (int s = 0; s < wb.NumberOfSheets; s++)
                {
                    var sheet = wb.GetSheetAt(s);
                    var sheetInfo = new BbsSheetInfo
                    {
                        SheetIndex = s,
                        SheetName  = sheet.SheetName
                    };

                    ScanSheet(sheet, sheetInfo);
                    info.Sheets.Add(sheetInfo);
                }
            }

            return info;
        }

        private static IWorkbook OpenWorkbook(Stream stream, string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".xls")  return new HSSFWorkbook(stream);
            if (ext == ".xlsx") return new XSSFWorkbook(stream);
            throw new NotSupportedException(
                "Only .xls and .xlsx are supported. Got: " + ext);
        }

        private static void ScanSheet(ISheet sheet, BbsSheetInfo info)
        {
            int lastRow = sheet.LastRowNum;

            // Krok 1: znajdź etykiety BOTTOM/TOP LAYER w kolumnie A
            //         + wiersz "Accessories" jako boundary końcowe.
            int? bottomLabelRow = null;
            int? topLabelRow    = null;
            for (int r = 0; r <= lastRow; r++)
            {
                string a = GetCellString(sheet, r, ColumnA);
                if (a == null) continue;
                string aUpper = a.Trim().ToUpperInvariant();

                if (aUpper == "BOTTOM LAYER" && !bottomLabelRow.HasValue)
                    bottomLabelRow = r;
                else if (aUpper == "TOP LAYER" && !topLabelRow.HasValue)
                    topLabelRow = r;
                else if (aUpper.StartsWith("ACCESSORIES")
                         && !info.AccessoriesRow.HasValue)
                    info.AccessoriesRow = r;
            }

            // Krok 2: dla każdej etykiety oblicz zakres jako PEŁNĄ POJEMNOŚĆ
            //         od label_row do (next_label - 1) lub (accessories - 1).
            //         NIE skanujemy kolumny B — działa na pustym template.
            int endBoundary;
            if (info.AccessoriesRow.HasValue)
                endBoundary = info.AccessoriesRow.Value;
            else
                endBoundary = lastRow + 1;  // brak accessories — bierzemy cały sheet

            if (bottomLabelRow.HasValue)
            {
                int first = bottomLabelRow.Value;
                int last;
                if (topLabelRow.HasValue && topLabelRow.Value > first)
                    last = topLabelRow.Value - 1;
                else
                    last = endBoundary - 1;

                if (last >= first)
                {
                    info.BottomLayer = new BbsLayerSection
                    {
                        LabelRow     = first,
                        FirstDataRow = first,
                        LastDataRow  = last
                    };
                }
            }

            if (topLabelRow.HasValue)
            {
                int first = topLabelRow.Value;
                int last  = endBoundary - 1;

                if (last >= first)
                {
                    info.TopLayer = new BbsLayerSection
                    {
                        LabelRow     = first,
                        FirstDataRow = first,
                        LastDataRow  = last
                    };
                }
            }
        }

        private static string GetCellString(ISheet sheet, int row, int col)
        {
            var r = sheet.GetRow(row);
            if (r == null) return null;
            var c = r.GetCell(col);
            if (c == null) return null;
            switch (c.CellType)
            {
                case CellType.String:  return c.StringCellValue;
                case CellType.Numeric:
                    return c.NumericCellValue.ToString(
                        System.Globalization.CultureInfo.InvariantCulture);
                case CellType.Boolean: return c.BooleanCellValue.ToString();
                case CellType.Formula:
                    try { return c.StringCellValue; }
                    catch
                    {
                        try { return c.NumericCellValue.ToString(); }
                        catch { return null; }
                    }
                default: return null;
            }
        }
    }
}
