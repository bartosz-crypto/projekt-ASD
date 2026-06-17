using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System.Collections.Generic;
using System.Linq;

namespace AsdRcSlab
{
    /// <summary>
    /// ASD-PRG (v5, p149): czyszczenie przez API <c>AcadDocument.PurgeAll()</c>
    /// (COM, przez <c>dynamic</c>) — KONIEC z makrem -PURGE w linii poleceń.
    ///
    /// Karmienie komendy -PURGE przez SendStringToExecute było narowiste: po
    /// zatwierdzeniu "All" komenda kończyła się, a zostawione tokeny ("* N")
    /// powtarzały -PURGE i zawieszały na "Enter name(s) to purge". PurgeAll() to
    /// ten sam bezpieczny natywny purge (respektuje referencje, NIE psuje
    /// layoutu/plotu), ale BEZ linii poleceń, bez promptów, SYNCHRONICZNIE.
    ///
    /// ZERO managed <c>db.Purge</c>, ZERO <c>Erase</c>, ZERO SendStringToExecute.
    /// Liczenie before/after jest WYŁĄCZNIE do odczytu → raport dokładny (nie approx).
    /// </summary>
    public static class PurgeCleanupService
    {
        // Stempel wersji — wypisywany na starcie komendy, żeby user wiedział, że
        // załadowała się WŁAŚCIWA (nowa) DLL, a nie stara kopia z innego bundla.
        public const string BuildStamp = "p149";

        /// <summary>
        /// Uruchamia bezpieczny natywny purge przez COM PurgeAll() (3 przebiegi =
        /// kaskada zagnieżdżonych). Zwraca raport tekstowy (różnice before/after
        /// per tablica symboli). ZERO SendStringToExecute/db.Purge/Erase.
        /// </summary>
        public static string RunNativePurge(Document doc)
        {
            if (doc == null) return "No active document.";
            var db = doc.Database;

            var before = CountSymbols(db);

            dynamic acadDoc = doc.GetAcadDocument();   // COM IAcadDocument
            for (int i = 0; i < 3; i++)                // 3 przebiegi: kaskada
            {
                try { acadDoc.PurgeAll(); }            // bezpieczny purge, synchronicznie
                catch { break; }
            }

            var after = CountSymbols(db);
            return BuildDiffReport(before, after);
        }

        // Stała kolejność wyświetlania kategorii.
        private static readonly string[] Order =
        {
            "Layers", "Blocks", "Linetypes", "Text styles",
            "Dim styles", "RegApps", "UCS", "Views"
        };

        // read-only liczenie rekordów tablic symboli (ŻADNEGO db.Purge ani Erase).
        private static Dictionary<string, int> CountSymbols(Database db)
        {
            var d = new Dictionary<string, int>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                int C(ObjectId tblId)
                {
                    int n = 0;
                    var t = (SymbolTable)tr.GetObject(tblId, OpenMode.ForRead);
                    foreach (ObjectId id in t) n++;
                    return n;
                }
                d["Layers"]      = C(db.LayerTableId);
                d["Blocks"]      = C(db.BlockTableId);
                d["Linetypes"]   = C(db.LinetypeTableId);
                d["Text styles"] = C(db.TextStyleTableId);
                d["Dim styles"]  = C(db.DimStyleTableId);
                d["RegApps"]     = C(db.RegAppTableId);
                d["UCS"]         = C(db.UcsTableId);
                d["Views"]       = C(db.ViewTableId);
                tr.Commit();
            }
            return d;
        }

        private static string BuildDiffReport(
            Dictionary<string, int> before, Dictionary<string, int> after)
        {
            var lines = new List<string>();
            int total = 0;
            foreach (var cat in Order)
            {
                int diff = before[cat] - after[cat];
                if (diff > 0) { lines.Add($"  {cat}: -{diff}"); total += diff; }
            }
            if (total == 0) return "Nothing purged (drawing already clean).";
            lines.Add($"  TOTAL: -{total}");
            return string.Join("\n", lines);
        }
    }
}
