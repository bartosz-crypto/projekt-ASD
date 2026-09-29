using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AsdRcSlab
{
    public static class BbsXlsxWriter
    {
        public static void Write(string outputPath, List<BbsBarRow> rows)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path empty.", "outputPath");
            if (rows == null)
                throw new ArgumentNullException("rows");

            if (File.Exists(outputPath))
                File.Delete(outputPath);

            using (var package = new ExcelPackage(new FileInfo(outputPath)))
            {
                var sheet = package.Workbook.Worksheets.Add("BBS_BS8666");

                string[] headers = new[]
                {
                    "Bar mark", "Type and size", "No. of members",
                    "No. of bars in each", "Total no.", "Length per bar (raw)",
                    "Shape code", "A (mm)", "B (mm)", "C (mm)", "D (mm)",
                    "E/R (mm)", "BS8666 raw", "BS8666 final", "Status"
                };
                for (int c = 0; c < headers.Length; c++)
                {
                    var cell = sheet.Cells[1, c + 1];
                    cell.Value = headers[c];
                    cell.Style.Font.Bold = true;
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                for (int i = 0; i < rows.Count; i++)
                {
                    int r = i + 2;
                    var row = rows[i];
                    sheet.Cells[r, 1].Value = row.BarMark;
                    sheet.Cells[r, 2].Value = row.TypeSize;
                    sheet.Cells[r, 3].Value = row.NoMembers;
                    sheet.Cells[r, 4].Value = row.NoEach;
                    sheet.Cells[r, 5].Value = row.Total;
                    sheet.Cells[r, 6].Value = row.LengthPerBar;
                    sheet.Cells[r, 7].Value = row.ShapeCode;
                    sheet.Cells[r, 8].Value = row.A;
                    sheet.Cells[r, 9].Value = row.B;
                    sheet.Cells[r, 10].Value = row.C;
                    sheet.Cells[r, 11].Value = row.D;
                    if (row.EOrR.HasValue)
                        sheet.Cells[r, 12].Value = row.EOrR.Value;

                    double raw   = BS8666Calculator.CalculateRawCuttingLength(row);
                    double final = BS8666Calculator.CalculateFinalCuttingLength(row);

                    if (!double.IsNaN(raw))
                        sheet.Cells[r, 13].Value = Math.Round(raw, 2);
                    if (!double.IsNaN(final))
                        sheet.Cells[r, 14].Value = final;

                    sheet.Cells[r, 15].Value = BuildStatus(row, raw);
                }

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                package.Save();
            }
        }

        private static string BuildStatus(BbsBarRow row, double raw)
        {
            if (!double.IsNaN(raw)) return "OK";
            if (row.ShapeCode == 12 || row.ShapeCode == 25
                || row.ShapeCode == 34 || row.ShapeCode == 35
                || row.ShapeCode == 41 || row.ShapeCode == 44
                || row.ShapeCode == 46 || row.ShapeCode == 56)
            {
                if (!row.EOrR.HasValue)
                    return "ERROR: E/R required for shape code " + row.ShapeCode;
            }
            return "ERROR: unsupported shape code " + row.ShapeCode;
        }
    }
}
