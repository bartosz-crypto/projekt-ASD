using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AsdRcSlab
{
    /// <summary>
    /// Raport z operacji explode obiektów ASD (per warstwa + globalnie).
    /// </summary>
    public class ExplodeReport
    {
        public int LayersProcessed;          // ile warstw przekazano do explode
        public int TopLevelFound;            // ile encji top-level znaleziono
        public int ExplodedOk;               // ile encji rozbito poprawnie (top + nested)
        public int Failed;                   // ile rzuciło wyjątek przy Explode
        public int Skipped;                  // ile zwróciło 0 produktów (proxy/nierozbijalne)
        public int Created;                  // ile nowych prymitywów dodano
        public int LayersUnlocked;           // ile warstw odblokowano/odmrożono/włączono

        public List<string> NotExploded   = new List<string>();
        public List<string> UnlockedLayers = new List<string>();

        public Dictionary<string, int> PerLayerFound =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, int> PerLayerExploded =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Rozbija (EXPLODE) obiekty utworzone przez AutoCAD Structural Detailing,
    /// rozpoznawane po prefiksie warstwy "AutoCAD_Structural_Detailing_".
    /// Zwykłe obiekty AutoCAD (inne warstwy) nie są ruszane. Po rozbiciu zostaje
    /// czysta geometria, a oryginalne custom-obiekty ASD są kasowane.
    /// Operuje na całym Model Space.
    /// </summary>
    public class ExplodeAsdService
    {
        public const string AsdLayerPrefix = "AutoCAD_Structural_Detailing_";

        private static readonly string DiagLogPath =
            Path.Combine(
                Environment.GetEnvironmentVariable("TEMP") ?? @"C:\Temp",
                "AsdRcSlab-explode-diag.log");

        private static void Diag(string msg)
        {
            try
            {
                File.AppendAllText(DiagLogPath,
                    $"{DateTime.Now:HH:mm:ss.fff}  [EXPLODE] {msg}{Environment.NewLine}");
            }
            catch { /* ignore */ }
        }

        /// <summary>
        /// Faza A — SKAN (read-only). Zwraca mapę: warstwa ASD -> liczba encji top-level.
        /// </summary>
        public List<(string Layer, int Count)> ScanAsdLayers(Document doc)
        {
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (doc == null) return new List<(string, int)>();

            var db = doc.Database;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(
                    bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;
                    string layer = ent.Layer;
                    if (!string.IsNullOrEmpty(layer) &&
                        layer.StartsWith(AsdLayerPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        Inc(counts, layer);
                    }
                }
                tr.Commit();
            }

            return counts
                .OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase)
                .Select(k => (k.Key, k.Value))
                .ToList();
        }

        /// <summary>
        /// Faza B — EXPLODE (write). Wszystko w jednej transakcji + blokada dokumentu.
        /// </summary>
        public ExplodeReport ExplodeLayers(Document doc, IEnumerable<string> layers, bool recursive)
        {
            var rep = new ExplodeReport();
            if (doc == null) return rep;

            var selected = new HashSet<string>(
                layers ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            rep.LayersProcessed = selected.Count;
            Diag($"=== ExplodeLayers START === layers={selected.Count} recursive={recursive}");

            if (selected.Count == 0)
            {
                Diag("no layers selected - abort");
                return rep;
            }

            var db = doc.Database;
            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // 1. Odblokuj / odmróź / włącz zaznaczone warstwy ASD (jeśli trzeba).
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId ltrId in lt)
                {
                    var ltr = (LayerTableRecord)tr.GetObject(ltrId, OpenMode.ForRead);
                    if (!selected.Contains(ltr.Name)) continue;
                    if (!(ltr.IsLocked || ltr.IsFrozen || ltr.IsOff)) continue;

                    ltr.UpgradeOpen();
                    var states = new List<string>();
                    if (ltr.IsLocked) { ltr.IsLocked = false; states.Add("unlock"); }
                    if (ltr.IsFrozen) { ltr.IsFrozen = false; states.Add("thaw"); }
                    if (ltr.IsOff)    { ltr.IsOff    = false; states.Add("on"); }

                    rep.LayersUnlocked++;
                    rep.UnlockedLayers.Add($"{ltr.Name} ({string.Join("+", states)})");
                    Diag($"layer adjusted: {ltr.Name} -> {string.Join("+", states)}");
                }

                // 2. Zbierz NAJPIERW pełną listę encji top-level z zaznaczonych warstw.
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(
                    bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var targets = new List<ObjectId>();
                foreach (ObjectId id in ms)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;
                    if (!string.IsNullOrEmpty(ent.Layer) && selected.Contains(ent.Layer))
                        targets.Add(id);
                }
                rep.TopLevelFound = targets.Count;
                Diag($"top-level targets collected: {targets.Count}");

                // 3. Rozbij każdą encję top-level.
                foreach (ObjectId id in targets)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;
                    string layer = ent.Layer;
                    Inc(rep.PerLayerFound, layer);

                    bool ok = ExplodeRecursive(ent, ms, tr, recursive, rep, 0);
                    if (ok) Inc(rep.PerLayerExploded, layer);
                }

                tr.Commit();
            }

            Diag($"=== DONE === found={rep.TopLevelFound} ok={rep.ExplodedOk} " +
                 $"failed={rep.Failed} skipped={rep.Skipped} created={rep.Created} " +
                 $"unlocked={rep.LayersUnlocked}");
            return rep;
        }

        // Rozbija pojedynczą encję; rekursja TYLKO dla custom ASD (RBCR*/RBCT*).
        // Zwraca true gdy ta encja została rozbita i skasowana.
        private bool ExplodeRecursive(Entity ent, BlockTableRecord ms, Transaction tr,
                                      bool recursive, ExplodeReport rep, int depth)
        {
            var products = new DBObjectCollection();
            try
            {
                ent.Explode(products);            // NIE dodaje produktów do bazy
            }
            catch (System.Exception ex)
            {
                rep.Failed++;
                rep.NotExploded.Add(ent.Layer + " : " + ex.Message);
                Diag($"FAILED explode layer={ent.Layer} depth={depth}: {ex.Message}");
                return false;                     // oryginału NIE kasuj
            }

            if (products.Count == 0)              // proxy / nierozbijalny
            {
                rep.Skipped++;
                Diag($"SKIPPED (0 products) layer={ent.Layer} depth={depth}");
                return false;                     // oryginału NIE kasuj
            }

            foreach (DBObject obj in products)
            {
                var child = obj as Entity;
                if (child == null) { obj.Dispose(); continue; }

                ms.AppendEntity(child);
                tr.AddNewlyCreatedDBObject(child, true);
                rep.Created++;

                // Rekurencja tylko dla kontenerów ASD, żeby NIE szatkować
                // zwykłej geometrii (linii/polilinii/tekstu) w nieskończoność.
                string cls = child.GetRXClass().Name;
                bool isAsdCustom =
                    cls.StartsWith("RBCR", StringComparison.OrdinalIgnoreCase) ||
                    cls.StartsWith("RBCT", StringComparison.OrdinalIgnoreCase);
                if (recursive && isAsdCustom && depth < 8)
                    ExplodeRecursive(child, ms, tr, recursive, rep, depth + 1);
            }

            // oryginał rozbity OK -> skasuj
            if (!ent.IsWriteEnabled) ent.UpgradeOpen();
            ent.Erase();
            rep.ExplodedOk++;
            return true;
        }

        private static void Inc(Dictionary<string, int> d, string key)
        {
            if (string.IsNullOrEmpty(key)) return;
            if (d.ContainsKey(key)) d[key]++;
            else d[key] = 1;
        }
    }
}
