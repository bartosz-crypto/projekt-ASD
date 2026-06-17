using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AsdRcSlab
{
    public class PurgeReport
    {
        // Kategoria -> liczba.
        public Dictionary<string, int> Counts =
            new Dictionary<string, int>(StringComparer.Ordinal);

        public int Total => Counts.Values.Sum();

        public void Inc(string cat, int n = 1)
        {
            if (Counts.ContainsKey(cat)) Counts[cat] += n;
            else Counts[cat] = n;
        }

        // Stała kolejność wyświetlania (tylko kategorie z >0).
        // p142: wyłącznie tablice symboli (native -PURGE robi resztę bezpiecznie).
        public static readonly string[] Order =
        {
            "Layers", "Blocks", "Linetypes", "Text styles", "Dim styles",
            "RegApps", "UCS", "Views"
        };

        public string BuildSummary()
        {
            var lines = new List<string>();
            foreach (var cat in Order)
                if (Counts.TryGetValue(cat, out int n) && n > 0)
                    lines.Add($"  {cat}: {n}");
            return string.Join(Environment.NewLine, lines);
        }
    }

    /// <summary>
    /// ASD-PRG (v3, p142): czyszczenie rysunku DELEGOWANE do natywnej komendy
    /// AutoCAD <c>-PURGE</c>, która respektuje wszystkie referencje (viewporty,
    /// layouty, plot styles) i nigdy nie psuje rysunku.
    ///
    /// Wcześniejszy ręczny Database.Purge + Erase (oraz kasowanie geometrii
    /// zerowej / pustego tekstu) usuwał obiekty wciąż referencjonowane przez
    /// nadpisania w viewportach / layouty → Access Violation 0x0050 przy wejściu
    /// w layout / plotowaniu. Tu NIE MA żadnego ręcznego Erase.
    ///
    /// <see cref="EstimateUnused"/> = read-only estymata (Database.Purge na
    /// kolekcji NICZEGO nie kasuje — tylko filtruje). <see cref="ApplyNativePurge"/>
    /// uruchamia natywny -PURGE i raportuje rzeczywiste różnice PRZED/PO.
    /// </summary>
    public class PurgeCleanupService
    {
        // Tablice symboli liczone do estymaty / raportu (bezpieczne).
        private static readonly (Func<Database, ObjectId> TableId, string Cat)[] SymbolTables =
        {
            (db => db.LayerTableId,     "Layers"),
            (db => db.BlockTableId,     "Blocks"),
            (db => db.LinetypeTableId,  "Linetypes"),
            (db => db.TextStyleTableId, "Text styles"),
            (db => db.DimStyleTableId,  "Dim styles"),
            (db => db.RegAppTableId,    "RegApps"),
            (db => db.UcsTableId,       "UCS"),
            (db => db.ViewTableId,      "Views"),
        };

        private static readonly string DiagLogPath =
            Path.Combine(
                Environment.GetEnvironmentVariable("TEMP") ?? @"C:\Temp",
                "AsdRcSlab-purge-diag.log");

        private static void Diag(string msg)
        {
            try
            {
                File.AppendAllText(DiagLogPath,
                    $"{DateTime.Now:HH:mm:ss.fff}  [PURGE] {msg}{Environment.NewLine}");
            }
            catch { /* ignore */ }
        }

        private static string Fmt(PurgeReport rep) =>
            string.Join(", ", rep.Counts.Where(k => k.Value > 0).Select(k => $"{k.Key}={k.Value}"));

        /// <summary>
        /// READ-ONLY estymata nieużywanych obiektów nazwanych. Buduje kolekcję
        /// ID z tablic symboli i woła <c>db.Purge(coll)</c> — to TYLKO odfiltrowuje
        /// kolekcję (zostają purgeable), NIC nie kasuje w bazie. Liczby są
        /// orientacyjne (native -PURGE może usunąć nieco inny zestaw, kaskadowo).
        /// </summary>
        public PurgeReport EstimateUnused(Document doc)
        {
            var rep = new PurgeReport();
            if (doc == null) return rep;
            var db = doc.Database;

            Diag("=== Estimate START (read-only) ===");

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ids = new ObjectIdCollection();
                var catMap = new Dictionary<ObjectId, string>();

                foreach (var (tableId, cat) in SymbolTables)
                {
                    var tbl = tr.GetObject(tableId(db), OpenMode.ForRead) as SymbolTable;
                    if (tbl == null) continue;
                    foreach (ObjectId id in tbl) { ids.Add(id); catMap[id] = cat; }
                }

                db.Purge(ids);   // read-only: zostają tylko purgeable, nic nie kasuje
                foreach (ObjectId id in ids)
                    rep.Inc(catMap.TryGetValue(id, out var c) ? c : "Blocks");

                tr.Abort();      // dodatkowa gwarancja: żadnego zapisu
            }

            Diag($"=== Estimate DONE (read-only) === est total={rep.Total} [{Fmt(rep)}]");
            return rep;
        }

        /// <summary>
        /// APLIKACJA: uruchamia natywne <c>-PURGE</c> przez
        /// <c>doc.SendStringToExecute</c> (Editor.Command nie istnieje w tym API).
        /// 3 zagnieżdżone przebiegi — usunięcie bloku zwalnia warstwę/linetyp itd.
        /// ZERO ręcznego Erase — native -PURGE respektuje wszystkie referencje.
        ///
        /// UWAGA: SendStringToExecute jest ASYNCHRONICZNE — komenda wykona się po
        /// zakończeniu bieżącej. Dlatego rzeczywistych liczb PO nie da się policzyć
        /// tu synchronicznie; do raportu używamy estymaty z <see cref="EstimateUnused"/>
        /// (orientacyjnej).
        /// </summary>
        public void QueueNativePurge(Document doc)
        {
            if (doc == null) return;

            Diag("=== Native -PURGE QUEUED (async SendStringToExecute, 3 passes) ===");

            // Tokeny "_"-prefixed = niezależne od języka. "_All"=wszystkie typy,
            // "*"=wszystkie nazwy, "_No"=bez pytań o każdy obiekt.
            const string purge = "_-PURGE _All * _No ";
            doc.SendStringToExecute(purge + purge + purge, true, false, false);
        }
    }
}
