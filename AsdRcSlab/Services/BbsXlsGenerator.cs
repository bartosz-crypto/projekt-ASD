using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace AsdRcSlab
{
    public sealed class BbsGenerateResult
    {
        public bool   Success          { get; set; }
        public string Message          { get; set; }
        public int    BottomPages      { get; set; }
        public int    TopPages         { get; set; }
        public int    BottomBarsWritten { get; set; }
        public int    TopBarsWritten   { get; set; }
        public string OutputPath       { get; set; }
    }

    public static class BbsXlsGenerator
    {
        // Wiersze (0-indexed) w template — zgodne z BBS layoutem:
        private const int RowA3 = 2;   // description
        private const int RowB5 = 4;   // Contract No. value
        private const int RowH5 = 4;   // Address line 1
        private const int RowH6 = 5;   // Address line 2
        private const int RowH7 = 6;   // Address line 3
        private const int RowB7 = 6;   // Revision value
        private const int RowL8 = 7;   // Page X of Y
        private const int RowM5 = 4;   // Drg list

        // Kolumny (0-indexed):
        private const int ColA = 0;
        private const int ColB = 1;
        private const int ColH = 7;
        private const int ColL = 11;
        private const int ColM = 12;

        // Accessories block — wiersze 41+ (0-indexed: 40+).
        private const int AccessoriesStartRow = 40;
        private const int AccessoriesEndRow   = 60;

        // Ostatnia kolumna danych (M):
        private const int DataColLast = 12;

        /// <summary>
        /// Generuje nowy multi-page BBS na podstawie template + danych.
        /// </summary>
        public static BbsGenerateResult Generate(
            BbsGenerationContext context,
            List<BbsBarRow>      rows,
            string               templatePath,
            string               outputPath)
        {
            if (context == null)
                throw new ArgumentNullException("context");
            if (rows == null)
                throw new ArgumentNullException("rows");
            if (string.IsNullOrWhiteSpace(templatePath))
                throw new ArgumentException("Template path empty.");
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path empty.");
            if (!File.Exists(templatePath))
                throw new FileNotFoundException("Template not found.", templatePath);

            // 1. Wczytaj template
            IWorkbook wb;
            using (var fs = new FileStream(
                templatePath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite))
            {
                wb = OpenWorkbook(fs, templatePath);
            }

            // 2. Wykryj capacity w template (re-use BbsXlsReader)
            var layout = BbsXlsReader.DetectLayout(templatePath);
            if (layout.Sheets.Count == 0)
                return Fail("Template has no sheets.");
            var templateSheetInfo = layout.Sheets[0];
            if (templateSheetInfo.BottomLayer == null)
                return Fail(
                    "Template sheet must have BOTTOM LAYER label in column A. "
                    + "Plugin uses this section as base for cloning.");

            int capacity          = templateSheetInfo.BottomLayer.CapacityRows;
            int firstDataRow      = templateSheetInfo.BottomLayer.FirstDataRow;
            int templateSheetIndex = templateSheetInfo.SheetIndex;

            // 3. Podziel rows
            var bottom = rows.Where(r => r.BarMark < 100).ToList();
            var top    = rows.Where(r => r.BarMark >= 100).ToList();

            // 4. Walidacja
            if (bottom.Count > 0 && context.BottomLayouts.Count == 0)
                return Fail(
                    "Input has {0} bottom-layer bars but no layouts "
                    + "assigned as BOTTOM in dialog.", bottom.Count);
            if (top.Count > 0 && context.TopLayouts.Count == 0)
                return Fail(
                    "Input has {0} top-layer bars but no layouts "
                    + "assigned as TOP in dialog.", top.Count);

            // 5. Page counts
            int bottomPages = bottom.Count == 0 ? 0
                : (int)Math.Ceiling((double)bottom.Count / capacity);
            int topPages = top.Count == 0 ? 0
                : (int)Math.Ceiling((double)top.Count / capacity);

            if (bottomPages == 0 && topPages == 0)
                return Fail("No bars to write.");

            // 6a. Generate BOTTOM sheets
            string bottomDrgList = BuildDrgList(context.BottomLayouts);
            string bottomPrefix  = BuildSheetPrefix(
                context.BottomLayouts, context.Revision);
            int writtenBottom = 0;

            for (int p = 1; p <= bottomPages; p++)
            {
                ISheet newSheet = wb.CloneSheet(templateSheetIndex);
                string sheetName = BuildSheetName(bottomPrefix, p, bottomPages);
                int newIdx = wb.GetSheetIndex(newSheet.SheetName);
                wb.SetSheetName(newIdx, sheetName);

                var pageBars = bottom
                    .Skip((p - 1) * capacity)
                    .Take(capacity)
                    .ToList();

                FillSheet(
                    newSheet, context, bottomDrgList,
                    description:  BuildA3("BOTTOM LAYER", context.PlotSuffix),
                    layerLabel:   "BOTTOM LAYER",
                    pageNumber:   p,
                    totalPages:   bottomPages,
                    firstDataRow: firstDataRow,
                    capacity:     capacity,
                    pageBars:     pageBars);

                if (p > 1) ClearAccessories(newSheet);

                writtenBottom += pageBars.Count;
            }

            // 6b. Generate TOP sheets
            string topDrgList = BuildDrgList(context.TopLayouts);
            string topPrefix  = BuildSheetPrefix(
                context.TopLayouts, context.Revision);
            int writtenTop = 0;

            for (int p = 1; p <= topPages; p++)
            {
                ISheet newSheet = wb.CloneSheet(templateSheetIndex);
                string sheetName = BuildSheetName(topPrefix, p, topPages);
                int newIdx = wb.GetSheetIndex(newSheet.SheetName);
                wb.SetSheetName(newIdx, sheetName);

                var pageBars = top
                    .Skip((p - 1) * capacity)
                    .Take(capacity)
                    .ToList();

                FillSheet(
                    newSheet, context, topDrgList,
                    description:  BuildA3("TOP LAYER", context.PlotSuffix),
                    layerLabel:   "TOP LAYER",
                    pageNumber:   p,
                    totalPages:   topPages,
                    firstDataRow: firstDataRow,
                    capacity:     capacity,
                    pageBars:     pageBars);

                if (p > 1) ClearAccessories(newSheet);

                writtenTop += pageBars.Count;
            }

            // 7. Usuń oryginalny template sheet
            wb.RemoveSheetAt(templateSheetIndex);

            // 8. Zapisz
            if (File.Exists(outputPath)) File.Delete(outputPath);
            using (var fs = new FileStream(
                outputPath, FileMode.Create, FileAccess.Write))
            {
                wb.Write(fs);
            }
            wb.Close();

            return new BbsGenerateResult
            {
                Success          = true,
                BottomPages      = bottomPages,
                TopPages         = topPages,
                BottomBarsWritten = writtenBottom,
                TopBarsWritten   = writtenTop,
                OutputPath       = outputPath,
                Message          = string.Format(
                    "Generated {0} BOTTOM page(s) ({1} bars) + "
                    + "{2} TOP page(s) ({3} bars).",
                    bottomPages, writtenBottom, topPages, writtenTop)
            };
        }

        // --- helpers ---

        private static BbsGenerateResult Fail(
            string fmt, params object[] args)
        {
            return new BbsGenerateResult
            {
                Success = false,
                Message = "ABORT: " + string.Format(fmt, args)
            };
        }

        private static IWorkbook OpenWorkbook(Stream s, string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".xls")  return new HSSFWorkbook(s);
            if (ext == ".xlsx") return new XSSFWorkbook(s);
            throw new NotSupportedException("Unsupported ext: " + ext);
        }

        /// <summary>
        /// Buduje listę rysunków rozdzielonych przecinkami.
        /// Bierze człon po ostatnim "-" w DrawingNumber.
        /// ["RH149ZS001-RC010","RH149ZS001-RC012"] → "RC010, RC012"
        /// </summary>
        private static string BuildDrgList(List<BbsLayoutInfo> layouts)
        {
            var parts = new List<string>();
            foreach (var lay in layouts)
            {
                string drg = lay.DrawingNumber ?? lay.LayoutName ?? "";
                int dash = drg.LastIndexOf('-');
                if (dash >= 0 && dash < drg.Length - 1)
                    drg = drg.Substring(dash + 1);
                parts.Add(drg);
            }
            return string.Join(", ", parts);
        }

        /// <summary>
        /// Prefix dla sheet name: "RC010-RC012-RC013-C1".
        /// </summary>
        private static string BuildSheetPrefix(
            List<BbsLayoutInfo> layouts, string revision)
        {
            var parts = new List<string>();
            foreach (var lay in layouts)
            {
                string drg = lay.DrawingNumber ?? lay.LayoutName ?? "";
                int dash = drg.LastIndexOf('-');
                if (dash >= 0 && dash < drg.Length - 1)
                    drg = drg.Substring(dash + 1);
                parts.Add(drg);
            }
            string list = string.Join("-", parts);
            string rev  = string.IsNullOrWhiteSpace(revision)
                ? "" : "-" + revision;
            return list + rev;
        }

        private static string BuildSheetName(
            string prefix, int page, int totalPages)
        {
            if (totalPages <= 1) return prefix;
            return string.Format(
                "{0}-Page-{1}of{2}", prefix, page, totalPages);
        }

        private static string BuildA3(string layerLabel, string plotSuffix)
        {
            const string baseDesc =
                "REINFORCEMENT DETAILS OF SPEEDECK PILED RAFT FOUNDATION - ";
            if (string.IsNullOrWhiteSpace(plotSuffix))
                return baseDesc + layerLabel;
            return baseDesc + layerLabel + " - " + plotSuffix;
        }

        /// <summary>
        /// Wypełnia jeden arkusz: title block + label + bary.
        /// </summary>
        private static void FillSheet(
            ISheet                sheet,
            BbsGenerationContext  context,
            string                drgList,
            string                description,
            string                layerLabel,
            int                   pageNumber,
            int                   totalPages,
            int                   firstDataRow,
            int                   capacity,
            List<BbsBarRow>       pageBars)
        {
            // A3 description
            SetString(sheet, RowA3, ColA, description);

            // B5 Contract No.
            SetString(sheet, RowB5, ColB, context.ContractNo ?? "");

            // H5/H6/H7 Address
            SetString(sheet, RowH5, ColH, context.AddressLine1 ?? "");
            SetString(sheet, RowH6, ColH, context.AddressLine2 ?? "");
            SetString(sheet, RowH7, ColH, context.AddressLine3 ?? "");

            // B7 Revision
            SetString(sheet, RowB7, ColB, context.Revision ?? "");

            // M5 Drg list
            SetString(sheet, RowM5, ColM, drgList);

            // L8 Page X of Y
            SetString(sheet, RowL8, ColL,
                string.Format("Page {0} of {1}", pageNumber, totalPages));

            // A15 (firstDataRow) layer label
            SetString(sheet, firstDataRow, ColA, layerLabel);

            // Dane barów — reuse BbsXlsWriter.WriteOneRowPublic
            for (int i = 0; i < pageBars.Count; i++)
            {
                int r   = firstDataRow + i;
                var row = sheet.GetRow(r) ?? sheet.CreateRow(r);
                BbsXlsWriter.WriteOneRowPublic(row, pageBars[i]);
            }
        }

        private static void SetString(
            ISheet sheet, int rowIdx, int colIdx, string value)
        {
            var row  = sheet.GetRow(rowIdx) ?? sheet.CreateRow(rowIdx);
            var cell = row.GetCell(colIdx)  ?? row.CreateCell(colIdx);
            cell.SetCellValue(value);
        }

        private static void ClearAccessories(ISheet sheet)
        {
            for (int r = AccessoriesStartRow; r <= AccessoriesEndRow; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;
                for (int c = 0; c <= DataColLast; c++)
                {
                    var cell = row.GetCell(c);
                    if (cell == null) continue;
                    if (cell.CellType != CellType.Blank)
                        cell.SetCellType(CellType.Blank);
                }
            }
        }
    }
}
