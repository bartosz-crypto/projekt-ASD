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
        private const int ColumnB = 1;        // Bar mark

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
                else if (aUpper.StartsWith("ACCESSORIES") && !info.AccessoriesRow.HasValue)
                    info.AccessoriesRow = r;
            }

            int boundary = info.AccessoriesRow ?? lastRow;

            if (bottomLabelRow.HasValue)
            {
                int firstBottom = bottomLabelRow.Value;
                int lastBottom = FindLastDataRow(
                    sheet, firstBottom,
                    topLabelRow.HasValue ? topLabelRow.Value - 1 : boundary - 1);
                if (lastBottom >= firstBottom)
                {
                    info.BottomLayer = new BbsLayerSection
                    {
                        LabelRow     = firstBottom,
                        FirstDataRow = firstBottom,
                        LastDataRow  = lastBottom
                    };
                }
            }

            if (topLabelRow.HasValue)
            {
                int firstTop = topLabelRow.Value;
                int lastTop  = FindLastDataRow(sheet, firstTop, boundary - 1);
                if (lastTop >= firstTop)
                {
                    info.TopLayer = new BbsLayerSection
                    {
                        LabelRow     = firstTop,
                        FirstDataRow = firstTop,
                        LastDataRow  = lastTop
                    };
                }
            }
        }

        private static int FindLastDataRow(ISheet sheet, int firstRow, int maxRow)
        {
            int last = firstRow - 1;
            for (int r = firstRow; r <= maxRow; r++)
            {
                string b = GetCellString(sheet, r, ColumnB);
                if (!string.IsNullOrWhiteSpace(b)) last = r;
            }
            return last;
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
