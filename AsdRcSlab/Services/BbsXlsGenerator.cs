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
        public bool   Success           { get; set; }
        public string Message           { get; set; }
        public int    BottomPages       { get; set; }
        public int    TopPages          { get; set; }
        public int    BottomBarsWritten { get; set; }
        public int    TopBarsWritten    { get; set; }
        public string OutputPath        { get; set; }
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
        private const int RowM5 = 4;   // Drg list line 1 (alias dla czytelności)

        // Kolumny (0-indexed):
        private const int ColA = 0;
        private const int ColB = 1;
        private const int ColH = 7;
        private const int ColI = 8;
        private const int ColL = 11;
        private const int ColM = 12;

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

            int capacity           = templateSheetInfo.BottomLayer.CapacityRows;
            int firstDataRow       = templateSheetInfo.BottomLayer.FirstDataRow;
            int templateSheetIndex = templateSheetInfo.SheetIndex;

            // Referencja do template sheet — potrzebna do CopyColumnWidths
            var templateSheet = wb.GetSheetAt(templateSheetIndex);

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

            // Tylko PIERWSZY wygenerowany sheet zachowuje pełne accessories
            bool isFirstSheetInFile = true;

            // 6a. Generate BOTTOM sheets
            var bottomDrgLines = BuildDrgListLines(context.BottomLayouts);
            string bottomPrefix = BuildSheetPrefix(
                context.BottomLayouts, context.Revision);
            int writtenBottom = 0;

            for (int p = 1; p <= bottomPages; p++)
            {
                ISheet newSheet  = wb.CloneSheet(templateSheetIndex);
                string sheetName = BuildSheetName(bottomPrefix, p, bottomPages);
                int newIdx = wb.GetSheetIndex(newSheet.SheetName);
                wb.SetSheetName(newIdx, sheetName);
                CopyColumnWidths(templateSheet, newSheet);

                var pageBars = bottom
                    .Skip((p - 1) * capacity)
                    .Take(capacity)
                    .ToList();

                FillSheet(
                    newSheet, context, bottomDrgLines,
                    description:  BuildA3("BOTTOM LAYER", context.PlotSuffix),
                    layerLabel:   "BOTTOM LAYER",
                    pageNumber:   p,
                    totalPages:   bottomPages,
                    firstDataRow: firstDataRow,
                    capacity:     capacity,
                    pageBars:     pageBars);

                if (!isFirstSheetInFile)
                    ClearAccessoriesPartial(newSheet);
                isFirstSheetInFile = false;

                writtenBottom += pageBars.Count;
            }

            // 6b. Generate TOP sheets
            var topDrgLines = BuildDrgListLines(context.TopLayouts);
            string topPrefix = BuildSheetPrefix(
                context.TopLayouts, context.Revision);
            int writtenTop = 0;

            for (int p = 1; p <= topPages; p++)
            {
                ISheet newSheet  = wb.CloneSheet(templateSheetIndex);
                string sheetName = BuildSheetName(topPrefix, p, topPages);
                int newIdx = wb.GetSheetIndex(newSheet.SheetName);
                wb.SetSheetName(newIdx, sheetName);
                CopyColumnWidths(templateSheet, newSheet);

                var pageBars = top
                    .Skip((p - 1) * capacity)
                    .Take(capacity)
                    .ToList();

                FillSheet(
                    newSheet, context, topDrgLines,
                    description:  BuildA3("TOP LAYER", context.PlotSuffix),
                    layerLabel:   "TOP LAYER",
                    pageNumber:   p,
                    totalPages:   topPages,
                    firstDataRow: firstDataRow,
                    capacity:     capacity,
                    pageBars:     pageBars);

                if (!isFirstSheetInFile)
                    ClearAccessoriesPartial(newSheet);
                isFirstSheetInFile = false;

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
                Success           = true,
                BottomPages       = bottomPages,
                TopPages          = topPages,
                BottomBarsWritten = writtenBottom,
                TopBarsWritten    = writtenTop,
                OutputPath        = outputPath,
                Message           = string.Format(
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
        /// Kopiuje szerokości kolumn z source sheet do destination sheet.
        /// NPOI HSSF CloneSheet nie zawsze kopiuje column widths poprawnie —
        /// efekt to "rozjechana" tabela w klonowanym arkuszu.
        /// </summary>
        private static void CopyColumnWidths(
            ISheet source, ISheet dest, int maxCol = 30)
        {
            for (int c = 0; c < maxCol; c++)
            {
                int width = source.GetColumnWidth(c);
                dest.SetColumnWidth(c, width);
            }
        }

        /// <summary>
        /// Dzieli listę layoutów na linie tekstu (max 2 layouty per linia).
        /// Linie nie-ostatnie kończą się przecinkiem.
        /// Przykłady:
        ///   [RC011]                      → ["RC011"]
        ///   [RC010, RC012]               → ["RC010, RC012"]
        ///   [RC010, RC012, RC013]        → ["RC010, RC012,", "RC013"]
        ///   [RC010, RC011, RC012, RC013] → ["RC010, RC011,", "RC012, RC013"]
        /// </summary>
        private static List<string> BuildDrgListLines(
            List<BbsLayoutInfo> layouts, int maxPerLine = 2)
        {
            var suffixes = new List<string>();
            foreach (var lay in layouts)
            {
                string drg = lay.DrawingNumber ?? lay.LayoutName ?? "";
                int dash = drg.LastIndexOf('-');
                if (dash >= 0 && dash < drg.Length - 1)
                    drg = drg.Substring(dash + 1);
                suffixes.Add(drg);
            }

            var lines = new List<string>();
            for (int i = 0; i < suffixes.Count; i += maxPerLine)
            {
                var slice  = suffixes.Skip(i).Take(maxPerLine).ToList();
                bool isLast = (i + maxPerLine >= suffixes.Count);
                string joined = string.Join(", ", slice);
                if (!isLast) joined += ",";
                lines.Add(joined);
            }
            return lines;
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
            ISheet               sheet,
            BbsGenerationContext context,
            List<string>         drgLines,
            string               description,
            string               layerLabel,
            int                  pageNumber,
            int                  totalPages,
            int                  firstDataRow,
            int                  capacity,
            List<BbsBarRow>      pageBars)
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

            // M5, M6, M7... — Drg list (jedna linia per komórka)
            for (int i = 0; i < drgLines.Count; i++)
                SetString(sheet, RowM5 + i, ColM, drgLines[i]);

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

        /// <summary>
        /// Czyści tylko fragmenty accessories block, zachowując layout
        /// (ramki, labele "Accessories", "Tonnage:", "Tying wire:" itd.).
        /// Stosowane na wszystkich arkuszach OPRÓCZ pierwszego w pliku.
        ///
        /// Usuwa:
        /// - Wiersz 42 (0-indexed 41): cała linia TRIC-TRAK
        /// - Cell I43 (0-indexed [42, ColI]): HYSTOOLS line description
        ///
        /// Zachowuje:
        /// - Wiersz 41: A41 "Accessories", K41 "Tonnage:", L41 wartość, M41 "t"
        /// - Wiersz 43: A43 "Tying wire:", B43 "Y", D43 "Spacers (50):",
        ///              F43 "bags", H43 "Deckchairs:"
        /// </summary>
        private static void ClearAccessoriesPartial(ISheet sheet)
        {
            // Wiersz 42 (0-indexed 41) — TRIC-TRAK: czyść wszystkie cells
            var row42 = sheet.GetRow(41);
            if (row42 != null)
            {
                for (int c = 0; c <= DataColLast; c++)
                {
                    var cell = row42.GetCell(c);
                    if (cell != null && cell.CellType != CellType.Blank)
                        cell.SetCellType(CellType.Blank);
                }
            }

            // Cell I43 (0-indexed [42, ColI]) — HYSTOOLS: tylko ta komórka
            var row43 = sheet.GetRow(42);
            if (row43 != null)
            {
                var i43 = row43.GetCell(ColI);
                if (i43 != null && i43.CellType != CellType.Blank)
                    i43.SetCellType(CellType.Blank);
            }
        }
    }
}
