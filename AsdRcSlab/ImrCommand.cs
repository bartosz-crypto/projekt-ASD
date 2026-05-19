using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AsdRcSlab
{
    public class ImrCommand
    {
        // Bbox jednej ramki PH-FRAME (closed LWPOLYLINE)
        public class FrameBbox
        {
            public double Xmin, Xmax, Ymin, Ymax;
            public ObjectId FrameId;
            public double Width  => Xmax - Xmin;
            public double Height => Ymax - Ymin;
        }

        // Info o pojedynczym plocie (4 ramki T1/T2/B1/B2 + label + ref point)
        public class PlotMapInfo
        {
            public string    Label;
            public FrameBbox T1, T2, B1, B2;
            public Point3d   ReferencePoint; // top-left T1
        }

        [CommandMethod("ASD-IMR")]
        public void CmdImportReinfMap()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var activeDb = doc.Database;

            // 1. OpenFileDialog (Microsoft.Win32, jak było)
            string path;
            var ofd = new Microsoft.Win32.OpenFileDialog
            {
                Filter     = "AutoCAD files (*.dxf;*.dwg)|*.dxf;*.dwg|DXF (*.dxf)|*.dxf|DWG (*.dwg)|*.dwg",
                Title      = "Select reinforcement maps source file",
                DefaultExt = ".dxf"
            };
            if (ofd.ShowDialog() != true)
            {
                ed.WriteMessage("\nImport cancelled.");
                return;
            }
            path = ofd.FileName;
            ed.WriteMessage($"\n[ASD-IMR] Source file: {path}");

            try
            {
                // 2. Otwórz sideDb i TRZYMAJ otwartą do końca (ObjectIds muszą być valid przy WblockCloneObjects)
                using (var sideDb = new Database(false, true))
                {
                    bool isDxf = path.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase);
                    if (isDxf) sideDb.DxfIn(path, null);
                    else       sideDb.ReadDwgFile(path, FileShare.Read, true, "");

                    // 3. Scan
                    var plots = ScanReinforcementMaps(sideDb);
                    if (plots.Count == 0)
                    {
                        ed.WriteMessage("\n[WARN] No reinforcement map plots found in selected file.");
                        ed.WriteMessage("\n       Expected: PH-SLAB-HEADER MTEXT + PH-FRAME polylines + PH-T1/T2/B1/B2-TITLE MTEXT/DBText.");
                        return;
                    }
                    ed.WriteMessage($"\nFound {plots.Count} plot(s) in source file.");

                    // 4. Dialog wyboru plotu (lub auto-select gdy 1)
                    PlotMapInfo plot;
                    if (plots.Count == 1)
                    {
                        plot = plots[0];
                        ed.WriteMessage($"\n[INFO] Only one plot found, auto-selected: {plot.Label}");
                    }
                    else
                    {
                        var dialog = new ImrPlotPickerDialog(plots);
                        AcApp.ShowModalWindow(AcApp.MainWindow.Handle, dialog, false);
                        if (dialog.SelectedPlot == null)
                        {
                            ed.WriteMessage("\nImport cancelled.");
                            return;
                        }
                        plot = dialog.SelectedPlot;
                        ed.WriteMessage($"\n[INFO] Selected plot: {plot.Label}");
                    }

                    // 5. Drag preview dla wyboru punktu insercji
                    var jig = new ImrInsertionJig(plot);
                    var jigResult = ed.Drag(jig);
                    if (jigResult.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\nImport cancelled.");
                        return;
                    }
                    Point3d insertionPoint = jig.CurrentPoint;

                    // 6. Copy 4 maps to active drawing
                    CopyMapsToActiveDrawing(sideDb, activeDb, plot, insertionPoint, ed);

                    ed.WriteMessage("\n[INFO] Reinforcement maps imported successfully.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[ERROR] Failed during import: {ex.Message}");
            }
        }

        // ── Scan helpers ──────────────────────────────────────────────────────

        private List<PlotMapInfo> ScanReinforcementMaps(Database sideDb)
        {
            var plots = new List<PlotMapInfo>();

            using (var tr = sideDb.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(sideDb.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // 1. Find headers (MTEXT on PH-SLAB-HEADER)
                var headers = new List<(Point3d Pos, string Label)>();
                foreach (ObjectId id in ms)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;
                    if (!string.Equals(ent.Layer, "PH-SLAB-HEADER", StringComparison.OrdinalIgnoreCase)) continue;
                    if (ent is MText mt)
                    {
                        string label = ExtractFirstLineFromMText(mt.Contents ?? "");
                        if (string.IsNullOrWhiteSpace(label)) label = "(no label)";
                        headers.Add((mt.Location, label));
                    }
                }

                // 2. Find PH-FRAME polylines
                var frames = new List<FrameBbox>();
                foreach (ObjectId id in ms)
                {
                    var pl = tr.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                    if (pl == null) continue;
                    if (!string.Equals(pl.Layer, "PH-FRAME", StringComparison.OrdinalIgnoreCase)) continue;
                    Extents3d ext;
                    try { ext = pl.GeometricExtents; }
                    catch { continue; }
                    frames.Add(new FrameBbox
                    {
                        Xmin = ext.MinPoint.X, Xmax = ext.MaxPoint.X,
                        Ymin = ext.MinPoint.Y, Ymax = ext.MaxPoint.Y,
                        FrameId = id
                    });
                }

                // 3. Group frames by column (xmin within tolerance)
                const double colTol = 10.0;
                var columns = new List<List<FrameBbox>>();
                foreach (var f in frames)
                {
                    var existing = columns.FirstOrDefault(c => Math.Abs(c[0].Xmin - f.Xmin) < colTol);
                    if (existing != null) existing.Add(f);
                    else columns.Add(new List<FrameBbox> { f });
                }

                // 4. Match each header to a column (header.X in [column.Xmin - 500, column.Xmax])
                foreach (var header in headers)
                {
                    var col = columns.FirstOrDefault(c =>
                        header.Pos.X >= c[0].Xmin - 500.0 &&
                        header.Pos.X <= c[0].Xmax);
                    if (col == null || col.Count < 4) continue;

                    var t1 = FindFrameWithTitleLayer(col, ms, tr, "PH-T1-TITLE");
                    var t2 = FindFrameWithTitleLayer(col, ms, tr, "PH-T2-TITLE");
                    var b1 = FindFrameWithTitleLayer(col, ms, tr, "PH-B1-TITLE");
                    var b2 = FindFrameWithTitleLayer(col, ms, tr, "PH-B2-TITLE");

                    if (t1 == null || t2 == null || b1 == null || b2 == null) continue;

                    plots.Add(new PlotMapInfo
                    {
                        Label = header.Label,
                        T1 = t1, T2 = t2, B1 = b1, B2 = b2,
                        ReferencePoint = new Point3d(t1.Xmin, t1.Ymax, 0)
                    });
                }

                tr.Commit();
            }

            return plots.OrderBy(p => p.ReferencePoint.X).ToList();
        }

        private void CopyMapsToActiveDrawing(Database sideDb, Database activeDb,
            PlotMapInfo plot, Point3d insertionPoint, Editor ed)
        {
            // 1. Zbierz ObjectIds entities których centroid leży w którejkolwiek z 4 ramek
            var idsToClone = new ObjectIdCollection();
            var frames = new[] { plot.T1, plot.T2, plot.B1, plot.B2 };

            int collected = 0, skipped = 0;
            using (var trSrc = sideDb.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)trSrc.GetObject(sideDb.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)trSrc.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    var ent = trSrc.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null) { skipped++; continue; }

                    if (string.Equals(ent.Layer, "PH-SLAB-HEADER", StringComparison.OrdinalIgnoreCase))
                    { skipped++; continue; }

                    Extents3d ext;
                    try { ext = ent.GeometricExtents; }
                    catch { skipped++; continue; }

                    double cx = (ext.MinPoint.X + ext.MaxPoint.X) / 2.0;
                    double cy = (ext.MinPoint.Y + ext.MaxPoint.Y) / 2.0;

                    bool inAnyFrame = false;
                    foreach (var f in frames)
                    {
                        if (cx >= f.Xmin && cx <= f.Xmax && cy >= f.Ymin && cy <= f.Ymax)
                        { inAnyFrame = true; break; }
                    }

                    if (inAnyFrame) { idsToClone.Add(id); collected++; }
                    else skipped++;
                }

                trSrc.Commit();
            }

            ed.WriteMessage($"\n[INFO] Collected {collected} entities from 4 frames (skipped {skipped}).");

            if (idsToClone.Count == 0)
            {
                ed.WriteMessage("\n[WARN] No entities to copy — check frame bbox detection.");
                return;
            }

            // 2. WblockCloneObjects — cross-database clone
            var idMap = new IdMapping();
            ObjectId msDestId = SymbolUtilityServices.GetBlockModelSpaceId(activeDb);
            sideDb.WblockCloneObjects(idsToClone, msDestId, idMap,
                DuplicateRecordCloning.Ignore, false);

            // 3. Transform sklonowane entities do insertion point
            Vector3d displacement = insertionPoint - plot.ReferencePoint;
            Matrix3d xform = Matrix3d.Displacement(displacement);

            int transformed = 0;
            using (var trDest = activeDb.TransactionManager.StartTransaction())
            {
                foreach (IdPair pair in idMap)
                {
                    if (!pair.IsPrimary) continue;
                    if (!pair.IsCloned) continue;

                    var ent = trDest.GetObject(pair.Value, OpenMode.ForWrite) as Entity;
                    if (ent != null) { ent.TransformBy(xform); transformed++; }
                }
                trDest.Commit();
            }

            ed.WriteMessage($"\n[INFO] Cloned {idsToClone.Count} entities, transformed {transformed}.");
        }

        private FrameBbox FindFrameWithTitleLayer(List<FrameBbox> column, BlockTableRecord ms,
            Transaction tr, string titleLayer)
        {
            foreach (ObjectId id in ms)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (!string.Equals(ent.Layer, titleLayer, StringComparison.OrdinalIgnoreCase)) continue;

                // Akceptuj zarówno MText jak i DBText (single-line TEXT)
                Point3d p;
                if (ent is MText mt)
                    p = mt.Location;
                else if (ent is DBText dbtxt)
                    p = dbtxt.Position;
                else
                    continue;

                var frame = column.FirstOrDefault(f =>
                    p.X >= f.Xmin && p.X <= f.Xmax &&
                    p.Y >= f.Ymin && p.Y <= f.Ymax);
                if (frame != null) return frame;
            }
            return null;
        }

        // Wyciąga pierwszą linię z MText.Contents pomijając formatting tags
        // (\C5;\LPLOT 4-5\l\P\C7;... → "PLOT 4-5")
        private static string ExtractFirstLineFromMText(string contents)
        {
            if (string.IsNullOrEmpty(contents)) return "";

            int pIdx = contents.IndexOf(@"\P", StringComparison.Ordinal);
            string firstPara = pIdx >= 0 ? contents.Substring(0, pIdx) : contents;

            // Krok 1: tagi z parametrami zakończone ; (np. \C5;  \H1.5x;  \fArial|b0|i0|c0|p34;)
            string clean = Regex.Replace(firstPara, @"\\[A-Za-z][^;\\]*;", "");

            // Krok 2: tagi bez parametrów — backslash + jedna litera (np. \L \l \O \o)
            clean = Regex.Replace(clean, @"\\[A-Za-z]", "");

            clean = clean.Replace("{", "").Replace("}", "");

            return clean.Trim();
        }

    }
}
