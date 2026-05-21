using System.Collections.Generic;

namespace AsdRcSlab
{
    /// <summary>
    /// Wykryta struktura pliku BBS .xls — lista arkuszy z opisem
    /// gdzie są sekcje BOTTOM LAYER / TOP LAYER, ile wierszy zajmują,
    /// gdzie kończy się obszar zbrojeniowy (boundary z "Accessories").
    /// </summary>
    public sealed class BbsXlsLayoutInfo
    {
        public string FilePath { get; set; }
        public List<BbsSheetInfo> Sheets { get; set; }
            = new List<BbsSheetInfo>();
    }

    public sealed class BbsSheetInfo
    {
        public int SheetIndex { get; set; }   // 0-based (NPOI convention)
        public string SheetName { get; set; }
        public BbsLayerSection BottomLayer { get; set; }   // null jeśli brak
        public BbsLayerSection TopLayer { get; set; }      // null jeśli brak
        public int? AccessoriesRow { get; set; }
            // 0-based numer wiersza gdzie A=="Accessories" (boundary)
    }

    public sealed class BbsLayerSection
    {
        public int LabelRow { get; set; }    // 0-based wiersz z etykietą
        public int FirstDataRow { get; set; }
            // = LabelRow (etykieta jest w tym samym wierszu co 1. wiersz danych)
        public int LastDataRow { get; set; }
            // 0-based ostatni wiersz z danymi (B != null)
        public int DataRowCount
        {
            get { return LastDataRow - FirstDataRow + 1; }
        }
    }
}
