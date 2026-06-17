using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.IO;

namespace AsdRcSlab
{
    /// <summary>
    /// ASD-PRG (v6, p150): czyszczenie przez NATYWNĄ komendę AutoCAD <c>-PURGE</c>
    /// dostarczoną PLIKIEM SKRYPTU (.scr) uruchamianym komendą <c>SCRIPT</c>.
    ///
    /// DLACZEGO komenda, a nie API: walidacja headless (accoreconsole, AutoCAD 2015)
    /// potwierdziła, że komenda `-PURGE All * N` (×3) NIE wprowadza błędów/dangling
    /// (AUDIT po purge = baseline drawing). Programowe API (Database.Purge+Erase,
    /// COM PurgeAll) crashowało wejście w layout/plot (Access Violation 0x0050) —
    /// usuwało obiekty referencjonowane przez layouty/plot, których komenda pilnuje.
    ///
    /// DLACZEGO SCRIPT, a nie surowe SendStringToExecute(-PURGE…): przez SCRIPT
    /// procesor skryptu karmi prompty komendy linia-po-linii z poszanowaniem granic
    /// komend → nie ma efektu „resztkowy Enter powtarza -PURGE" (który wieszał makro
    /// na "Enter name(s) to purge"). Walidacja headless: 3 bloki skonsumowane czysto.
    ///
    /// ZERO db.Purge / Erase / COM PurgeAll. Liczenie before/after = WYŁĄCZNIE
    /// odczyt (purge jest asynchroniczny → after liczony w CommandEnded po SCRIPT).
    /// </summary>
    public static class PurgeCleanupService
    {
        // Stempel wersji — wypisywany na starcie komendy, żeby user wiedział, że
        // załadowała się WŁAŚCIWA (nowa) DLL, a nie stara kopia z innego bundla.
        public const string BuildStamp = "p151";

        private static readonly string ScriptPath =
            Path.Combine(Path.GetTempPath(), "AsdRcSlab_purge.scr");

        // 3 bloki = kaskada (usunięcie bloku zwalnia warstwę/linetyp w kolejnym).
        // Każdy blok: -PURGE↵ All↵ *↵ N↵ (verify=No). p151: ostatnie dwie linie
        // przywracają FILEDIA=1 (ustawiamy 0 z API przed SCRIPT, by nie było okna
        // wyboru pliku). BEZ pustych linii / QUIT (pusta linia = Enter = powtórzenie
        // komendy; QUIT zamknąłby AutoCAD). Plik NIE kończy się pustą linią.
        private const string PurgeBlock = "-PURGE\r\nAll\r\n*\r\nN\r\n";
        private const string ScriptBody = PurgeBlock + PurgeBlock + PurgeBlock + "FILEDIA\r\n1\r\n";

        /// <summary>
        /// Liczy „before", zapisuje .scr, rejestruje jednorazowy CommandEnded
        /// (po SCRIPT liczy „after" i wypisuje różnice), po czym uruchamia SCRIPT
        /// przez SendStringToExecute (FILEDIA 0/1 = bez okna wyboru pliku).
        /// Zwraca status startowy (purge jest asynchroniczny — raport per kategoria
        /// pojawi się na linii poleceń po zakończeniu skryptu).
        /// </summary>
        public static string RunNativePurge(Document doc)
        {
            if (doc == null) return "No active document.";
            var db = doc.Database;
            var ed = doc.Editor;

            var before = CountSymbols(db);

            try { File.WriteAllText(ScriptPath, ScriptBody); }
            catch (Exception ex) { return "Failed to write purge script: " + ex.Message; }

            // Jednorazowy hook: po zakończeniu komendy SCRIPT policz „after" i raport.
            CommandEventHandler handler = null;
            handler = (s, e) =>
            {
                if (!string.Equals(e.GlobalCommandName, "SCRIPT", StringComparison.OrdinalIgnoreCase))
                    return;                          // ignoruj zagnieżdżone -PURGE
                doc.CommandEnded -= handler;
                try
                {
                    var after = CountSymbols(db);
                    ed.WriteMessage("\nPRG: done.\n" + BuildDiffReport(before, after) + "\n");
                }
                catch { /* report best-effort */ }
            };
            doc.CommandEnded += handler;

            // p151: FILEDIA ustaw z API (NIE przez makro — makrowe "FILEDIA 0 …"
            // rozjeżdżało tokeny i zacinało na "Enter new value for FILEDIA"). Makro
            // = TYLKO SCRIPT ze ścieżką (ukośniki w przód, brak kłopotu z '\"').
            // FILEDIA=1 przywraca ostatnia linia w .scr (skrypt wykonuje liniowo).
            string p = ScriptPath.Replace("\\", "/");
            Application.SetSystemVariable("FILEDIA", 0);
            doc.SendStringToExecute("_SCRIPT \"" + p + "\"\n", true, false, false);

            return "Purge started (native -PURGE x3 via SCRIPT). Per-category report follows on the command line.";
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
