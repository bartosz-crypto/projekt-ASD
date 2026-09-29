using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace AsdRcSlab
{
    public static class BbsXlsxReader
    {
        public static List<BbsBarRow> Read(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path is empty.", "path");
            if (!File.Exists(path))
                throw new FileNotFoundException("File not found.", path);

            var result = new List<BbsBarRow>();

            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                    throw new InvalidOperationException(
                        "Workbook has no worksheets.");

                var sheet = package.Workbook.Worksheets[1];  // EPPlus 4.x: 1-indexed

                // Auto-detect format:
                //  - calculated.xlsx (z ASD-BBC):
                //      A1="Bar mark" (string), dane od row 2
                //  - b1.xlsx (raw z ASD):
                //      A1=puste lub inne, dane od row 4
                int firstDataRow = DetectFirstDataRow(sheet);

                int row = firstDataRow;
                while (row <= sheet.Dimension.End.Row)
                {
                    var markRaw = sheet.Cells[row, 1].Value;
                    if (markRaw == null
                        || string.IsNullOrWhiteSpace(markRaw.ToString()))
                    {
                        row++;
                        continue;
                    }

                    // Skip wierszy gdzie BarMark NIE jest liczbą (header rows
                    // typu "Bar mark" jeśli auto-detect dał false positive).
                    if (!IsNumeric(markRaw))
                    {
                        row++;
                        continue;
                    }

                    var barRow = new BbsBarRow
                    {
                        BarMark      = ParseInt(markRaw),
                        TypeSize     = ParseString(sheet.Cells[row, 2].Value),
                        NoMembers    = ParseInt(sheet.Cells[row, 3].Value),
                        NoEach       = ParseInt(sheet.Cells[row, 4].Value),
                        Total        = ParseInt(sheet.Cells[row, 5].Value),
                        LengthPerBar = ParseDouble(sheet.Cells[row, 6].Value),
                        ShapeCode    = ParseInt(sheet.Cells[row, 7].Value),
                        A            = ParseDouble(sheet.Cells[row, 8].Value),
                        B            = ParseDouble(sheet.Cells[row, 9].Value),
                        C            = ParseDouble(sheet.Cells[row, 10].Value),
                        D            = ParseDouble(sheet.Cells[row, 11].Value),
                        EOrR         = ParseDoubleNullable(sheet.Cells[row, 12].Value)
                    };

                    result.Add(barRow);
                    row++;
                }
            }

            return result;
        }

        /// <summary>
        /// Detekcja formatu po nagłówkach:
        /// - A1 = "Bar mark" (lub coś z "Bar mark" prefix) → calculated.xlsx,
        ///   dane od row 2.
        /// - Inaczej (puste albo "SPEEDECK FOUNDATIONS" itp.) → b1-style,
        ///   dane od row 4.
        /// </summary>
        private static int DetectFirstDataRow(
            OfficeOpenXml.ExcelWorksheet sheet)
        {
            var a1 = sheet.Cells[1, 1].Value;
            if (a1 == null) return 4;  // pusty → b1-style
            string a1Text = a1.ToString().Trim();
            if (a1Text.StartsWith(
                "Bar mark", System.StringComparison.OrdinalIgnoreCase))
                return 2;
            return 4;
        }

        private static bool IsNumeric(object value)
        {
            if (value == null) return false;
            string s = value.ToString().Trim();
            if (string.IsNullOrEmpty(s)) return false;
            double d;
            return double.TryParse(s, out d);
        }

        private static int ParseInt(object value)
        {
            if (value == null) return 0;
            int n;
            if (int.TryParse(value.ToString(), out n)) return n;
            double d;
            if (double.TryParse(value.ToString(), out d)) return (int)Math.Round(d);
            return 0;
        }

        private static double ParseDouble(object value)
        {
            if (value == null) return 0.0;
            double d;
            return double.TryParse(value.ToString(), out d) ? d : 0.0;
        }

        private static double? ParseDoubleNullable(object value)
        {
            if (value == null) return null;
            string s = value.ToString();
            if (string.IsNullOrWhiteSpace(s)) return null;
            double d;
            return double.TryParse(s, out d) ? (double?)d : null;
        }

        private static string ParseString(object value)
        {
            return value == null ? "" : value.ToString().Trim();
        }
    }
}
