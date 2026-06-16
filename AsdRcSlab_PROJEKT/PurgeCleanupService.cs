using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AsdRcSlab
{
    public class PurgeReport
    {
        // Kategoria -> liczba usuniętych.
        public Dictionary<string, int> Counts =
            new Dictionary<string, int>(StringComparer.Ordinal);

        public List<string> DictsUsed = new List<string>();     // słowniki obecne w bazie
        public List<string> DictsSkipped = new List<string>();  // brak w API/bazie 2015

        // Obiekty zwrócone przez Purge jako purgeable, ale nieusuwalne (chronione
        // domyślne AutoCAD: eVSIsAcadDefault, domyślne materiały itd.).
        public int SkippedProtected;

        public int Total => Counts.Values.Sum();   // FAKTYCZNIE usunięte (bez protected)

        public void Inc(string cat, int n = 1)
        {
            if (Counts.ContainsKey(cat)) Counts[cat] += n;
            else Counts[cat] = n;
        }

        // Stała kolejność wyświetlania (tylko kategorie z >0).
        public static readonly string[] Order =
        {
            "Layers", "Blocks", "Linetypes", "Text styles", "Dim styles",
            "Table styles", "MLine styles", "MLeader styles", "Visual styles",
            "Materials", "Groups", "Plot styles", "RegApps", "UCS", "Views",
            "Zero-length geometry", "Empty text"
        };

        public string BuildSummary()
        {
            var lines = new List<string>();
            foreach (var cat in Order)
                if (Counts.TryGetValue(cat, out int n) && n > 0)
                    lines.Add($"  {cat}: {n}");
            if (SkippedProtected > 0)
                lines.Add($"  Skipped (protected): {SkippedProtected}");
            return string.Join(Environment.NewLine, lines);
        }
    }

    /// <summary>
    /// ASD-PRG: pełne czyszczenie rysunku — usuwa NIEUŻYWANE obiekty nazwane
    /// (warstwy, bloki, linetypy, style…) przez zarządzane Database.Purge oraz
    /// geometrię zerową i pusty tekst. commit=false → dry-run (Abort, dokładne
    /// liczby). commit=true → faktyczne usunięcie (Commit).
    /// </summary>
    public class PurgeCleanupService
    {
        private const double Eps = 1e-6;

        // Słowniki w Named Objects Dictionary (wersjoodporne — bez zależności od
        // db.XxxDictionaryId, których część istnieje dopiero w API 2017+).
        private static readonly (string Key, string Cat)[] DictSources =
        {
            ("ACAD_MLINESTYLE",      "MLine styles"),
            ("ACAD_TABLESTYLE",      "Table styles"),
            ("ACAD_MLEADERSTYLE",    "MLeader styles"),
            ("ACAD_VISUALSTYLE",     "Visual styles"),
            ("ACAD_MATERIAL",        "Materials"),
            ("ACAD_GROUP",           "Groups"),
            ("ACAD_PLOTSTYLENAME",   "Plot styles"),
            ("ACAD_DETAILVIEWSTYLE", "Detail view styles"),
            ("ACAD_SECTIONVIEWSTYLE","Section view styles"),
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

        public PurgeReport Run(Document doc, bool commit)
        {
            var rep = new PurgeReport();
            if (doc == null) return rep;
            var db = doc.Database;

            Diag($"=== Run START === commit={commit}");

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // 1. Geometria zerowa + pusty tekst (najpierw — zwalnia referencje).
                PurgeBadGeometry(db, tr, rep);

                // 2. Obiekty nazwane — iteracyjny Database.Purge aż do wyczerpania.
                PurgeNamedObjects(db, tr, rep);

                if (commit) tr.Commit(); else tr.Abort();
            }

            Diag($"=== Run DONE === commit={commit} total={rep.Total} " +
                 $"[{string.Join(", ", rep.Counts.Where(k => k.Value > 0).Select(k => $"{k.Key}={k.Value}"))}]");
            return rep;
        }

        // ---------- obiekty nazwane ----------
        private void PurgeNamedObjects(Database db, Transaction tr, PurgeReport rep)
        {
            var catMap = new Dictionary<ObjectId, string>();
            var remaining = new HashSet<ObjectId>();

            void AddTable(ObjectId tableId, string cat)
            {
                var tbl = tr.GetObject(tableId, OpenMode.ForRead) as SymbolTable;
                if (tbl == null) return;
                foreach (ObjectId id in tbl)
                {
                    if (remaining.Add(id)) catMap[id] = cat;
                }
            }

            AddTable(db.LayerTableId,     "Layers");
            AddTable(db.BlockTableId,     "Blocks");
            AddTable(db.LinetypeTableId,  "Linetypes");
            AddTable(db.TextStyleTableId, "Text styles");
            AddTable(db.DimStyleTableId,  "Dim styles");
            AddTable(db.RegAppTableId,    "RegApps");
            AddTable(db.UcsTableId,       "UCS");
            AddTable(db.ViewTableId,      "Views");

            // Słowniki z NOD (guard każdego).
            var nod = tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;
            foreach (var (key, cat) in DictSources)
            {
                try
                {
                    if (nod == null || !nod.Contains(key)) { rep.DictsSkipped.Add(cat); continue; }
                    var dict = tr.GetObject(nod.GetAt(key), OpenMode.ForRead) as DBDictionary;
                    if (dict == null) { rep.DictsSkipped.Add(cat); continue; }
                    rep.DictsUsed.Add(cat);
                    foreach (DBDictionaryEntry e in dict)
                    {
                        if (remaining.Add(e.Value)) catMap[e.Value] = cat;
                    }
                }
                catch (System.Exception ex)
                {
                    rep.DictsSkipped.Add(cat);
                    Diag($"dict {cat} ({key}) skipped: {ex.Message}");
                }
            }

            // Iteracyjny purge: Database.Purge zostawia w kolekcji tylko purgeable;
            // erase ich może zwolnić kolejne → powtarzaj aż pusto.
            int pass = 0;
            while (remaining.Count > 0)
            {
                pass++;
                var col = new ObjectIdCollection(remaining.ToArray());
                db.Purge(col);              // filtruje: zostają tylko purgeable
                if (col.Count == 0) break;

                int erasedThisPass = 0;
                foreach (ObjectId id in col)
                {
                    try
                    {
                        var obj = tr.GetObject(id, OpenMode.ForWrite, false, true);
                        if (obj == null) { remaining.Remove(id); continue; }
                        obj.Erase();
                        rep.Inc(catMap.TryGetValue(id, out var c) ? c : "Blocks");
                        erasedThisPass++;
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception)
                    {
                        // np. eVSIsAcadDefault — chroniony default, pomiń.
                        rep.SkippedProtected++;
                    }
                    catch (System.Exception)
                    {
                        rep.SkippedProtected++;
                    }
                    finally
                    {
                        // Zawsze usuń z puli — protected nie wracają w kolejnym przebiegu.
                        remaining.Remove(id);
                    }
                }

                // Progress-break: zostały same nieusuwalne defaulty → koniec.
                if (erasedThisPass == 0) break;
            }
            Diag($"named purge passes={pass} skippedProtected={rep.SkippedProtected}");
        }

        // ---------- geometria zerowa + pusty tekst ----------
        private void PurgeBadGeometry(Database db, Transaction tr, PurgeReport rep)
        {
            var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            if (bt == null) return;

            var toErase = new List<(ObjectId Id, string Cat)>();

            foreach (ObjectId btrId in bt)
            {
                var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null || !btr.IsLayout) continue;   // tylko Model + Paper space

                foreach (ObjectId entId in btr)
                {
                    var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;

                    string cat = ClassifyBad(ent);
                    if (cat != null) toErase.Add((entId, cat));
                }
            }

            int erased = 0;
            foreach (var (id, cat) in toErase)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForWrite, false, true) as Entity;
                    if (ent == null) continue;
                    ent.Erase();
                    rep.Inc(cat);
                    erased++;
                }
                catch (Autodesk.AutoCAD.Runtime.Exception) { rep.SkippedProtected++; }
                catch (System.Exception)                   { rep.SkippedProtected++; }
            }
            Diag($"bad geometry erased={erased} (candidates={toErase.Count})");
        }

        // Zwraca kategorię ("Zero-length geometry"/"Empty text") lub null.
        private static string ClassifyBad(Entity ent)
        {
            if (ent is DBText t)
                return string.IsNullOrWhiteSpace(t.TextString) ? "Empty text" : null;

            if (ent is MText mt)
            {
                string txt = mt.Text;
                if (string.IsNullOrWhiteSpace(txt)) txt = mt.Contents;
                return string.IsNullOrWhiteSpace(txt) ? "Empty text" : null;
            }

            if (ent is Line ln)
                return ln.StartPoint.DistanceTo(ln.EndPoint) < Eps ? "Zero-length geometry" : null;

            if (ent is Arc ar)
                return ar.Radius < Eps ? "Zero-length geometry" : null;

            if (ent is Circle ci)
                return ci.Radius < Eps ? "Zero-length geometry" : null;

            if (ent is Polyline pl)
                return (pl.NumberOfVertices < 2 || pl.Length < Eps) ? "Zero-length geometry" : null;

            if (ent is Polyline2d p2)
                return IsZeroOldPolyline(p2) ? "Zero-length geometry" : null;

            if (ent is Polyline3d p3)
                return IsZeroOldPolyline(p3) ? "Zero-length geometry" : null;

            return null;
        }

        private static bool IsZeroOldPolyline(Curve poly)
        {
            int vc = 0;
            if (poly is System.Collections.IEnumerable en)
                foreach (var _ in en) vc++;
            if (vc < 2) return true;
            try
            {
                double len = poly.GetDistanceAtParameter(poly.EndParam)
                           - poly.GetDistanceAtParameter(poly.StartParam);
                return len < Eps;
            }
            catch { return false; }
        }
    }
}
