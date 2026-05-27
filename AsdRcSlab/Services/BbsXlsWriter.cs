using System;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace AsdRcSlab
{
    /// <summary>
    /// Wynik operacji zapisu — dla raportowania user'owi.
    /// </summary>
    public sealed class BbsWriteResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public int BottomWritten { get; set; }
        public int TopWritten { get; set; }
        public int BottomCleared { get; set; }
        public int TopCleared { get; set; }
    }

    public static class BbsXlsWriter
    {
        // Stałe layout (0-indexed)
        private const int ColB = 1;   // Bar mark
        private const int ColC = 2;   // Type
        private const int ColD = 3;   // No. members
        private const int ColE = 4;   // No. each
        private const int ColF = 5;   // Total
        private const int ColG = 6;   // Length per bar (BS8666 final)
        private const int ColH = 7;   // Shape code
        private const int ColI = 8;   // A
        private const int ColJ = 9;   // B
        private const int ColK = 10;  // C
        private const int ColL = 11;  // D
        private const int ColM = 12;  // E/R

        /// <summary>
        /// Zapisuje listę BbsBarRow do BBS .xls. Bary z BarMark&lt;100 idą
        /// do BOTTOM, ≥100 do TOP. Wykorzystuje detekcję layoutu z
        /// BbsXlsReader. Jeśli rows nie mieszczą się w sekcji — ABORT
        /// (bez zapisu pliku).
        /// </summary>
        public static BbsWriteResult Write(
            string bbsPath, List<BbsBarRow> rows)
        {
            if (string.IsNullOrWhiteSpace(bbsPath))
                throw new ArgumentException("Path empty.", "bbsPath");
            if (!File.Exists(bbsPath))
                throw new FileNotFoundException("BBS not found.", bbsPath);
            if (rows == null)
                throw new ArgumentNullException("rows");

            // 1. Detekcja layoutu (z p100)
            var layout = BbsXlsReader.DetectLayout(bbsPath);

            // 2. Podział rows na bottom/top
            var bottom = new List<BbsBarRow>();
            var top    = new List<BbsBarRow>();
            foreach (var r in rows)
            {
                if (r.BarMark < 100) bottom.Add(r);
                else top.Add(r);
            }

            // 3. Znajdź sekcje BOTTOM i TOP (mogą być w różnych arkuszach!)
            BbsSheetInfo     bottomSheet = null;
            BbsLayerSection  bottomSec   = null;
            BbsSheetInfo     topSheet    = null;
            BbsLayerSection  topSec      = null;
            foreach (var s in layout.Sheets)
            {
                if (s.BottomLayer != null && bottomSec == null)
                {
                    bottomSheet = s;
                    bottomSec   = s.BottomLayer;
                }
                if (s.TopLayer != null && topSec == null)
                {
                    topSheet = s;
                    topSec   = s.TopLayer;
                }
            }

            // 4. Walidacja: czy mamy gdzie pisać + czy mieści się
            if (bottom.Count > 0 && bottomSec == null)
                return Abort(
                    "Found {0} bottom-layer bars in input but BBS has no "
                    + "BOTTOM LAYER section.", bottom.Count);
            if (top.Count > 0 && topSec == null)
                return Abort(
                    "Found {0} top-layer bars in input but BBS has no "
                    + "TOP LAYER section.", top.Count);

            if (bottomSec != null && bottom.Count > bottomSec.CapacityRows)
                return Abort(
                    "BOTTOM overflow: {0} bars to write but section has "
                    + "only {1} rows (sheet '{2}', rows {3}-{4}). "
                    + "Auto-extend layout will be added later (p102+).",
                    bottom.Count, bottomSec.CapacityRows,
                    bottomSheet.SheetName,
                    bottomSec.FirstDataRow + 1, bottomSec.LastDataRow + 1);
            if (topSec != null && top.Count > topSec.CapacityRows)
                return Abort(
                    "TOP overflow: {0} bars to write but section has "
                    + "only {1} rows (sheet '{2}', rows {3}-{4}). "
                    + "Auto-extend layout will be added later (p102+).",
                    top.Count, topSec.CapacityRows,
                    topSheet.SheetName,
                    topSec.FirstDataRow + 1, topSec.LastDataRow + 1);

            // 5. Otwórz workbook, clear + write, save
            IWorkbook wb;
            using (var fs = new FileStream(
                bbsPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wb = OpenWorkbook(fs, bbsPath);
            }

            var result = new BbsWriteResult { Success = true };

            if (bottomSec != null)
            {
                var sheet = wb.GetSheetAt(bottomSheet.SheetIndex);
                result.BottomCleared = ClearSection(sheet, bottomSec);
                result.BottomWritten = WriteRows(sheet, bottomSec, bottom);
            }
            if (topSec != null)
            {
                var sheet = wb.GetSheetAt(topSheet.SheetIndex);
                result.TopCleared = ClearSection(sheet, topSec);
                result.TopWritten = WriteRows(sheet, topSec, top);
            }

            // Save (overwrites istniejący plik — backup robi komenda)
            using (var fs = new FileStream(
                bbsPath, FileMode.Create, FileAccess.Write))
            {
                wb.Write(fs);
            }
            wb.Close();

            result.Message = string.Format(
                "Wrote {0} bottom + {1} top bars. "
                + "(Cleared {2} bottom + {3} top rows before write.)",
                result.BottomWritten, result.TopWritten,
                result.BottomCleared, result.TopCleared);
            return result;
        }

        // --- helpers ---

        private static BbsWriteResult Abort(string fmt, params object[] args)
        {
            return new BbsWriteResult
            {
                Success = false,
                Message = "ABORT: " + string.Format(fmt, args)
            };
        }

        private static IWorkbook OpenWorkbook(Stream stream, string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".xls")  return new HSSFWorkbook(stream);
            if (ext == ".xlsx") return new XSSFWorkbook(stream);
            throw new NotSupportedException("Unsupported ext: " + ext);
        }

        /// <summary>
        /// Czyści kolumny B-M w sekcji (FirstDataRow..LastDataRow włącznie),
        /// zostawia kolumnę A (etykieta warstwy) i kolumny AB/AC (formuły).
        /// SetCellType(Blank) zachowuje CellStyle (ramki, formatowanie).
        /// Zwraca liczbę wierszy w których coś było.
        /// </summary>
        private static int ClearSection(ISheet sheet, BbsLayerSection sec)
        {
            int cleared = 0;
            for (int r = sec.FirstDataRow; r <= sec.LastDataRow; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;
                bool hadSomething = false;
                for (int c = ColB; c <= ColM; c++)
                {
                    var cell = row.GetCell(c);
                    if (cell != null && cell.CellType != CellType.Blank)
                    {
                        hadSomething = true;
                        // SetCellType(Blank) zachowuje CellStyle (formatowanie),
                        // w przeciwieństwie do RemoveCell z p101c.
                        cell.SetCellType(CellType.Blank);
                    }
                }
                if (hadSomething) cleared++;
            }
            return cleared;
        }

        private static int WriteRows(
            ISheet sheet, BbsLayerSection sec, List<BbsBarRow> rows)
        {
            int written = 0;
            for (int i = 0; i < rows.Count; i++)
            {
                int r   = sec.FirstDataRow + i;
                var row = sheet.GetRow(r) ?? sheet.CreateRow(r);
                WriteOneRow(row, rows[i]);
                written++;
            }
            return written;
        }

        private static void WriteOneRow(IRow row, BbsBarRow b)
        {
            int  code       = b.ShapeCode;
            bool isStraight = (code == 0);
            bool isLink     = (code == 51 || code == 63);

            // Zawsze pisane: B-G (bar mark, type, no mbrs, no each, total, length)
            GetOrCreate(row, ColB).SetCellValue(b.BarMark);
            GetOrCreate(row, ColC).SetCellValue(b.TypeSize ?? "");
            GetOrCreate(row, ColD).SetCellValue(b.NoMembers);
            GetOrCreate(row, ColE).SetCellValue(b.NoEach);
            GetOrCreate(row, ColF).SetCellValue(b.Total);

            // G: BS8666 final length
            double final = BS8666Calculator.CalculateFinalCuttingLength(b);
            if (!double.IsNaN(final))
                GetOrCreate(row, ColG).SetCellValue(final);

            // H: shape code
            if (isStraight)
                GetOrCreate(row, ColH).SetCellValue("00");
            else
                GetOrCreate(row, ColH).SetCellValue((double)code);

            // I-M: per-row logic
            if (isStraight)
            {
                // Special case code 0: I="STR", J-M nietykane.
                // Cells pre-existing po Clear są Blank z zachowanym stylem.
                GetOrCreate(row, ColI).SetCellValue("STR");
            }
            else if (isLink)
            {
                // Closed links (51, 63): tylko A i B. C, D, E/R NIE pisane.
                // (Dla 51/63 wymuszone blank w C/D nawet jeśli input ma
                // niezerowe wartości — Excel template formuła nie bierze
                // tych pod uwagę dla zamkniętych strzemion.)
                GetOrCreate(row, ColI).SetCellValue(b.A);
                GetOrCreate(row, ColJ).SetCellValue(b.B);
                // K (C), L (D), M (E/R) — NIE tykamy.
            }
            else
            {
                // Wszystkie pozostałe shape codes: A, B, C, D, E/R z input.
                // BEZ per-code dimension logic — user'owa decyzja po teście
                // SP116QL001 (shape code 15 gubił B).
                GetOrCreate(row, ColI).SetCellValue(b.A);
                GetOrCreate(row, ColJ).SetCellValue(b.B);
                GetOrCreate(row, ColK).SetCellValue(b.C);
                GetOrCreate(row, ColL).SetCellValue(b.D);
                if (b.EOrR.HasValue)
                    GetOrCreate(row, ColM).SetCellValue(b.EOrR.Value);
            }
        }

        /// <summary>
        /// Public wrapper — używany przez BbsXlsGenerator aby uniknąć
        /// duplikacji logiki per-code dimensions z p101d.
        /// </summary>
        public static void WriteOneRowPublic(IRow row, BbsBarRow b)
        {
            WriteOneRow(row, b);
        }

        private static ICell GetOrCreate(IRow row, int col)
        {
            return row.GetCell(col) ?? row.CreateCell(col);
        }
    }
}
