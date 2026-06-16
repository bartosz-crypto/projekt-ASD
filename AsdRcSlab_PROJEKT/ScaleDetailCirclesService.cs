using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace AsdRcSlab
{
    /// <summary>
    /// Kandydat: ramka detalu 1:25 z kółkami rozkładu do przeskalowania.
    /// </summary>
    public class DetailCandidate
    {
        public ObjectId FrameId;
        public double W;
        public double H;
        public double CenterX;
        public double CenterY;
        public int Color;                       // efektywny ACI ramki
        public List<ObjectId> CircleIds = new List<ObjectId>();
        public int CircleCount;
        public string NearestLabelText = "";       // skrót "DETAIL 1:NN"
        public string NearestLabelFull = "";        // pełny oczyszczony tekst etykiety
        public double NearestLabelDist;
        public bool Preselected;
    }

    public class ScaleReport
    {
        public int FramesApproved;
        public int CirclesScaled;               // unikalne kółka przeskalowane
        public int CirclesSkipped;              // R < MIN (idempotencja)
        public int LayersUnlocked;
        public List<string> UnlockedLayers = new List<string>();
    }

    /// <summary>
    /// ASD-SDC: wykrywa ramki detali 1:25 (przerywane, kolor 1/10, zamknięte)
    /// zawierające kółka rozkładu (CIRCLE na warstwie "rozkład pręta_", R>=MIN)
    /// i skaluje te kółka ×0.5 (środek bez zmian). Skala 1:25 czytana z etykiety
    /// MTEXT "DETAIL 'XX' ... SCALE 1:25".
    /// </summary>
    public class ScaleDetailCirclesService
    {
        // ASCII-safe; trafia w "rozkład pręta_" (po prefiksie jest 'rozk').
        // NIE myl z "opis rozkładu_" (po prefiksie jest 'opis ' przed 'rozk').
        private const string RozkLayerMatch = "AutoCAD_Structural_Detailing_rozk";
        private const string OpisLayerPrefix = "AutoCAD_Structural_Detailing_opis";

        private static readonly HashSet<string> DashedLinetypes =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "DASHED", "DASHED2", "DASH", "HIDDEN", "HIDDEN2" };

        private static readonly HashSet<int> BorderColors = new HashSet<int> { 1, 10 };

        private const int TargetDetailScale = 25;     // 1:25
        private const double ScaleFactor = 0.5;        // o połowę
        private const double MinRadiusToScale = 25.0;  // 37.5 tak, 18.0/18.75 nie
        private const double PreselectLabelDist = 3000.0;

        private static readonly string DiagLogPath =
            Path.Combine(
                Environment.GetEnvironmentVariable("TEMP") ?? @"C:\Temp",
                "AsdRcSlab-scaledetail-diag.log");

        private static void Diag(string msg)
        {
            try
            {
                File.AppendAllText(DiagLogPath,
                    $"{DateTime.Now:HH:mm:ss.fff}  [SDC] {msg}{Environment.NewLine}");
            }
            catch { /* ignore */ }
        }

        // Usuwa kody formatowania MTEXT, żeby regex "SCALE 1:NN" działał także na
        // "SCALE 1:\Fromans|c238;25" -> "SCALE 1:25".
        public static string StripMTextCodes(string s)
        {
            if (string.IsNullOrEmpty(s)) return s ?? "";
            string t = s;

            // Stacking \S num/den; -> num/den (zachowaj tekst, usuń \S i ;)
            t = Regex.Replace(t, @"\\S([^;]*);", "$1");
            // Komendy z parametrem zakończonym ';' : \F..; \f..; \C..; \H..; \W..;
            //   \T..; \Q..; \p..; \A..;  (np. \Fromans|c238;)
            t = Regex.Replace(t, @"\\[A-Za-z][^;\\]*;", "");
            // Łamanie linii / akapit / twarda spacja
            t = t.Replace(@"\P", " ").Replace(@"\p", " ").Replace(@"\~", " ");
            // Przełączniki bez parametru: \L \l \O \o \K \k
            t = Regex.Replace(t, @"\\[LlOoKk]", "");
            // Escapowane znaki
            t = t.Replace(@"\{", "{").Replace(@"\}", "}").Replace(@"\\", "\\");
            // Pozostałe nawiasy grupujące
            t = t.Replace("{", "").Replace("}", "");
            // Zwiń wielokrotne białe znaki do pojedynczej spacji + Trim
            // (np. zlepione akapity "DETAIL '1'3No.SCALE" stają się czytelne).
            t = Regex.Replace(t, @"\s+", " ").Trim();
            return t;
        }

        private static readonly Regex ScaleRegex =
            new Regex(@"SCALE\s*1\s*:\s*(\d+)", RegexOptions.IgnoreCase);

        // ===================== Faza A — SKAN (read-only) =====================
        public List<DetailCandidate> ScanCandidates(Document doc)
        {
            var result = new List<DetailCandidate>();
            if (doc == null) return result;
            var db = doc.Database;

            var circles = new List<(ObjectId Id, Point2d C, double R)>();
            var labels = new List<(int Nn, Point2d P, string Full)>();
            var frames = new List<(ObjectId Id, Point2d[] Poly, Point2d Centroid,
                                   double W, double H, int Color)>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(
                    bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;

                    // --- CIRCLES (rozkład pręta_, R >= MIN) ---
                    if (ent is Circle circle)
                    {
                        string lay = circle.Layer ?? "";
                        if (lay.StartsWith(RozkLayerMatch, StringComparison.OrdinalIgnoreCase) &&
                            !lay.StartsWith(OpisLayerPrefix, StringComparison.OrdinalIgnoreCase) &&
                            circle.Radius >= MinRadiusToScale)
                        {
                            circles.Add((id,
                                new Point2d(circle.Center.X, circle.Center.Y),
                                circle.Radius));
                        }
                        continue;
                    }

                    // --- LABELS (MTEXT / DBText: DETAIL + SCALE 1:NN) ---
                    if (ent is MText mt)
                    {
                        TryAddLabel(StripMTextCodes(mt.Contents),
                            new Point2d(mt.Location.X, mt.Location.Y), labels);
                        continue;
                    }
                    if (ent is DBText dt)
                    {
                        TryAddLabel(StripMTextCodes(dt.TextString),
                            new Point2d(dt.Position.X, dt.Position.Y), labels);
                        continue;
                    }

                    // --- FRAMES (closed LWPOLYLINE, dashed, kolor 1/10) ---
                    if (ent is Polyline pl && pl.Closed && pl.NumberOfVertices >= 3)
                    {
                        int effColor = EffectiveColor(pl, tr);
                        if (!BorderColors.Contains(effColor)) continue;
                        if (!IsDashed(EffectiveLinetype(pl, tr))) continue;

                        int n = pl.NumberOfVertices;
                        var poly = new Point2d[n];
                        double minX = double.MaxValue, minY = double.MaxValue;
                        double maxX = double.MinValue, maxY = double.MinValue;
                        double sx = 0, sy = 0;
                        for (int i = 0; i < n; i++)
                        {
                            var p = pl.GetPoint2dAt(i);
                            poly[i] = p;
                            sx += p.X; sy += p.Y;
                            if (p.X < minX) minX = p.X; if (p.X > maxX) maxX = p.X;
                            if (p.Y < minY) minY = p.Y; if (p.Y > maxY) maxY = p.Y;
                        }
                        var centroid = new Point2d(sx / n, sy / n);
                        frames.Add((id, poly, centroid, maxX - minX, maxY - minY, effColor));
                    }
                }

                tr.Commit();
            }

            Diag($"scan: circles(R>=MIN)={circles.Count} labels(DETAIL 1:NN)={labels.Count} " +
                 $"frames(dashed col1/10 closed)={frames.Count}");

            // Budowa kandydatów.
            foreach (var f in frames)
            {
                var inside = new List<ObjectId>();
                foreach (var c in circles)
                    if (PointInPolygon(c.C, f.Poly)) inside.Add(c.Id);

                if (inside.Count == 0) continue;

                // Najbliższa etykieta DETAIL-scale do centroidu ramki.
                if (labels.Count == 0) continue;
                int bestNn = -1; double bestDist = double.MaxValue; string bestFull = "";
                foreach (var l in labels)
                {
                    double d = (l.P - f.Centroid).Length;
                    if (d < bestDist) { bestDist = d; bestNn = l.Nn; bestFull = l.Full; }
                }
                if (bestNn != TargetDetailScale) continue;   // najbliższa etykieta != 1:25

                result.Add(new DetailCandidate
                {
                    FrameId = f.Id,
                    W = f.W,
                    H = f.H,
                    CenterX = f.Centroid.X,
                    CenterY = f.Centroid.Y,
                    Color = f.Color,
                    CircleIds = inside,
                    CircleCount = inside.Count,
                    NearestLabelText = $"DETAIL 1:{bestNn}",
                    NearestLabelFull = bestFull,
                    NearestLabelDist = bestDist,
                    Preselected = bestDist <= PreselectLabelDist
                });
            }

            Diag($"scan: candidates={result.Count} preselected={result.Count(c => c.Preselected)}");
            return result;
        }

        private static void TryAddLabel(string text, Point2d pos,
                                        List<(int, Point2d, string)> labels)
        {
            if (string.IsNullOrEmpty(text)) return;
            if (text.IndexOf("DETAIL", StringComparison.OrdinalIgnoreCase) < 0) return;
            var m = ScaleRegex.Match(text);
            if (!m.Success) return;
            if (int.TryParse(m.Groups[1].Value, out int nn))
                labels.Add((nn, pos, text));   // text już oczyszczony przez StripMTextCodes
        }

        // ===================== Faza B — APLIKACJA (write) =====================
        public ScaleReport ApplyScaling(Document doc, IEnumerable<DetailCandidate> selected)
        {
            var rep = new ScaleReport();
            if (doc == null) return rep;

            var sel = (selected ?? Enumerable.Empty<DetailCandidate>()).ToList();
            rep.FramesApproved = sel.Count;

            // Unia ObjectId kółek (to samo kółko może leżeć w dwóch ramkach).
            var ids = new HashSet<ObjectId>();
            foreach (var c in sel)
                foreach (var oid in c.CircleIds)
                    ids.Add(oid);

            Diag($"=== ApplyScaling START === frames={sel.Count} uniqueCircles={ids.Count}");
            if (ids.Count == 0) return rep;

            var db = doc.Database;
            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Odblokuj / odmróź / włącz warstwę "rozkład pręta_".
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId ltrId in lt)
                {
                    var ltr = (LayerTableRecord)tr.GetObject(ltrId, OpenMode.ForRead);
                    if (!ltr.Name.StartsWith(RozkLayerMatch, StringComparison.OrdinalIgnoreCase) ||
                        ltr.Name.StartsWith(OpisLayerPrefix, StringComparison.OrdinalIgnoreCase))
                        continue;
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

                foreach (ObjectId id in ids)
                {
                    var circle = tr.GetObject(id, OpenMode.ForRead) as Circle;
                    if (circle == null) continue;
                    if (circle.Radius < MinRadiusToScale)   // bezpiecznik idempotencji
                    {
                        rep.CirclesSkipped++;
                        continue;
                    }
                    circle.UpgradeOpen();
                    circle.Radius = circle.Radius * ScaleFactor;   // środek bez zmian
                    rep.CirclesScaled++;
                }

                tr.Commit();
            }

            Diag($"=== DONE === framesApproved={rep.FramesApproved} scaled={rep.CirclesScaled} " +
                 $"skipped={rep.CirclesSkipped} layersUnlocked={rep.LayersUnlocked}");
            return rep;
        }

        // ===================== Helpers =====================
        private static int EffectiveColor(Entity ent, Transaction tr)
        {
            if (ent.ColorIndex == 256)   // ByLayer
            {
                var ltr = tr.GetObject(ent.LayerId, OpenMode.ForRead) as LayerTableRecord;
                if (ltr != null) return ltr.Color.ColorIndex;
            }
            return ent.ColorIndex;
        }

        private static string EffectiveLinetype(Entity ent, Transaction tr)
        {
            string lt = ent.Linetype ?? "";
            if (string.IsNullOrEmpty(lt) ||
                lt.Equals("ByLayer", StringComparison.OrdinalIgnoreCase))
            {
                var ltr = tr.GetObject(ent.LayerId, OpenMode.ForRead) as LayerTableRecord;
                if (ltr != null)
                {
                    var ltype = tr.GetObject(ltr.LinetypeObjectId, OpenMode.ForRead)
                                as LinetypeTableRecord;
                    if (ltype != null) return ltype.Name;
                }
            }
            return lt;
        }

        private static bool IsDashed(string ltname)
        {
            if (string.IsNullOrEmpty(ltname)) return false;
            string n = ltname.Trim();
            if (DashedLinetypes.Contains(n)) return true;
            // Generozyjnie: dowolny linetype z "DASH"/"HIDDEN" w nazwie.
            return n.IndexOf("DASH", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("HIDDEN", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // Ray casting point-in-polygon.
        private static bool PointInPolygon(Point2d pt, Point2d[] poly)
        {
            bool inside = false;
            int n = poly.Length;
            for (int i = 0, j = n - 1; i < n; j = i++)
            {
                double xi = poly[i].X, yi = poly[i].Y;
                double xj = poly[j].X, yj = poly[j].Y;
                bool intersect = ((yi > pt.Y) != (yj > pt.Y)) &&
                    (pt.X < (xj - xi) * (pt.Y - yi) / (yj - yi) + xi);
                if (intersect) inside = !inside;
            }
            return inside;
        }
    }
}
