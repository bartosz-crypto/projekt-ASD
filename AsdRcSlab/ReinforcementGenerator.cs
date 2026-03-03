using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AsdRcSlab
{
    public class ReinfGenResult
    {
        public int BarsDrawn { get; set; }
        /// <summary>X coords of horizontal bar (B1/T1) lap joint centres.</summary>
        public List<double> LapPositionsX { get; set; } = new List<double>();
        /// <summary>Y coords of vertical bar (B2/T2) lap joint centres.</summary>
        public List<double> LapPositionsY { get; set; } = new List<double>();
        public string Log   { get; set; } = "";
        public string Error { get; set; } = "";
    }

    public static class ReinforcementGenerator
    {
        private const string BarLayer  = "AutoCAD_Structural_Detailing_Bar distribution";
        private const string PileLayer = "SD-Pile";
        private const double MaxBarLen = 6000.0;
        private const double StepLen   = 250.0;   // round-up granularity & stagger step

        private static readonly System.Globalization.CultureInfo _inv =
            System.Globalization.CultureInfo.InvariantCulture;

        // ── Public API ────────────────────────────────────────────────────────────

        public static ReinfGenResult GenerateBottom(Database db, ObjectId slabId)
        {
            return Generate(db, slabId,
                spacing: 200.0, cover: 40.0, lap: 500.0,
                isTop: false, bottomLapsX: null, bottomLapsY: null);
        }

        public static ReinfGenResult GenerateTop(Database db, ObjectId slabId,
            List<double> bottomLapsX, List<double> bottomLapsY)
        {
            return Generate(db, slabId,
                spacing: 200.0, cover: 40.0, lap: 600.0,
                isTop: true,
                bottomLapsX: bottomLapsX ?? new List<double>(),
                bottomLapsY: bottomLapsY ?? new List<double>());
        }

        // ── ASD command-string API (COPY-based, no modal dialogs) ─────────────────

        /// <summary>
        /// Computes B1/B2 bar positions and places them using AutoCAD COPY command.
        /// Template bars are taken from the catalog registered by ASD-GSETUP.
        /// B1 = horizontal bars; B2 = vertical bars (COPYed then ROTATEd 90°).
        /// </summary>
        public static ReinfGenResult GenerateBottomAsd(Document doc, ObjectId slabId,
            Dictionary<int, Point3d> templateBars)
        {
            return GenerateAsd(doc, slabId,
                spacing: 200.0, cover: 40.0, lap: 500.0,
                isTop: false, templateBars: templateBars,
                bottomLapsX: null, bottomLapsY: null);
        }

        /// <summary>
        /// Computes T1/T2 bar positions and places them using AutoCAD COPY command.
        /// </summary>
        public static ReinfGenResult GenerateTopAsd(Document doc, ObjectId slabId,
            Dictionary<int, Point3d> templateBars,
            List<double> bottomLapsX, List<double> bottomLapsY)
        {
            return GenerateAsd(doc, slabId,
                spacing: 200.0, cover: 40.0, lap: 600.0,
                isTop: true, templateBars: templateBars,
                bottomLapsX: bottomLapsX ?? new List<double>(),
                bottomLapsY: bottomLapsY ?? new List<double>());
        }

        private static ReinfGenResult GenerateAsd(Document doc, ObjectId slabId,
            double spacing, double cover, double lap, bool isTop,
            Dictionary<int, Point3d> templateBars,
            List<double> bottomLapsX, List<double> bottomLapsY)
        {
            var result = new ReinfGenResult();
            var db = doc.Database;

            // ── 1. Read slab geometry (read-only transaction) ─────────────────────
            List<Point3d> vertices = null;
            List<(Point3d Center, double Radius)> piles = null;
            double xMin = 0, yMin = 0, xMax = 0, yMax = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var slabEnt = tr.GetObject(slabId, OpenMode.ForRead) as Entity;
                    if (!(slabEnt is Polyline slab))
                    {
                        result.Error = "Encja nie jest polilinią (LWPolyline).";
                        tr.Abort();
                        return result;
                    }
                    vertices = GetPolylineVertices(slab);
                    if (vertices.Count < 3)
                    {
                        result.Error = "Polilinia ma mniej niż 3 wierzchołki.";
                        tr.Abort();
                        return result;
                    }
                    GetSlabBounds(vertices, out xMin, out yMin, out xMax, out yMax);
                    var bt  = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(
                        bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                    piles = FindPilesInSlab(tr, btr, xMin, yMin, xMax, yMax, vertices);
                    tr.Commit();
                }
                catch (Exception ex)
                {
                    result.Error = ex.Message;
                    tr.Abort();
                    return result;
                }
            }

            // ── 2. Build COPY command string ──────────────────────────────────────
            var lapPosX = new List<double>();
            var lapPosY = new List<double>();
            var cmdSb   = new System.Text.StringBuilder();
            int barCount = 0;
            int skipped  = 0;

            // B1/T1: horizontal bars — DISTRIBUTION from (x1,y) to (x2,y)
            double y = yMin + cover;
            int rowIdx = 0;
            while (y <= yMax - cover)
            {
                var segs = IntersectHorizontal(vertices, y);
                foreach (var (xStart, xEnd) in segs)
                {
                    var bars = SplitSegment(xStart, xEnd, rowIdx, lap, isTop,
                        piles, isHoriz: true, perpCoord: y,
                        bottomLaps: isTop ? bottomLapsX : null);

                    foreach (var (x1, x2) in bars)
                    {
                        if (!TryGetTemplate(templateBars, x2 - x1, out int lenKey, out Point3d tb))
                        { skipped++; continue; }
                        AppendDistribution(cmdSb, tb, lenKey, x1, y, x2, y);
                        barCount++;
                    }
                    for (int i = 0; i < bars.Count - 1; i++)
                        lapPosX.Add((bars[i].b + bars[i + 1].a) / 2.0);
                }
                y += spacing;
                rowIdx++;
            }

            // B2/T2: vertical bars — DISTRIBUTION from (x,y1) to (x,y2)
            // ASD orients the bar along the distribution range direction automatically
            double x = xMin + cover;
            rowIdx = 0;
            while (x <= xMax - cover)
            {
                var segs = IntersectVertical(vertices, x);
                foreach (var (yStart, yEnd) in segs)
                {
                    var bars = SplitSegment(yStart, yEnd, rowIdx, lap, isTop,
                        piles, isHoriz: false, perpCoord: x,
                        bottomLaps: isTop ? bottomLapsY : null);

                    foreach (var (y1, y2) in bars)
                    {
                        if (!TryGetTemplate(templateBars, y2 - y1, out int lenKey, out Point3d tb))
                        { skipped++; continue; }
                        AppendDistribution(cmdSb, tb, lenKey, x, y1, x, y2);
                        barCount++;
                    }
                    for (int i = 0; i < bars.Count - 1; i++)
                        lapPosY.Add((bars[i].b + bars[i + 1].a) / 2.0);
                }
                x += spacing;
                rowIdx++;
            }

            // ── 3. Start dialog auto-closer then send commands ────────────────────
            // Auto-closer runs in background thread:
            //   "Reinforcement detailing"   → OK     (accepts distribution settings)
            //   "Reinforcement description" → Cancel  (skips description, bars placed)
            if (barCount > 0)
            {
                AsdDialogAutoCloser.Start(timeoutMs: 300_000); // 5 min for large slabs
                doc.SendStringToExecute(cmdSb.ToString(), true, false, false);
            }

            result.BarsDrawn     = barCount;
            result.LapPositionsX = lapPosX;
            result.LapPositionsY = lapPosY;
            result.Log = $"Wysłano {barCount} prętów ASD ({(isTop ? "T1/T2" : "B1/B2")})" +
                         (skipped > 0 ? $", pominięto {skipped} (brak szablonu w katalogu)." : ".");
            return result;
        }

        // ── Core Line-based generator (fallback / testing) ───────────────────────

        private static ReinfGenResult Generate(Database db, ObjectId slabId,
            double spacing, double cover, double lap, bool isTop,
            List<double> bottomLapsX, List<double> bottomLapsY)
        {
            var result = new ReinfGenResult();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var slabEnt = tr.GetObject(slabId, OpenMode.ForRead) as Entity;
                    if (!(slabEnt is Polyline slab))
                    {
                        result.Error = "Encja nie jest polilinią (LWPolyline).";
                        tr.Abort();
                        return result;
                    }

                    var bt  = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(
                        bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    EnsureLayer(tr, db, BarLayer, 5);

                    var vertices = GetPolylineVertices(slab);
                    if (vertices.Count < 3)
                    {
                        result.Error = "Polilinia ma mniej niż 3 wierzchołki.";
                        tr.Abort();
                        return result;
                    }

                    GetSlabBounds(vertices,
                        out double xMin2, out double yMin2,
                        out double xMax2, out double yMax2);

                    var piles2 = FindPilesInSlab(tr, btr, xMin2, yMin2, xMax2, yMax2, vertices);
                    var lapPosX = new List<double>();
                    var lapPosY = new List<double>();

                    double y2 = yMin2 + cover;
                    int rowIdx2 = 0;
                    while (y2 <= yMax2 - cover)
                    {
                        var segs = IntersectHorizontal(vertices, y2);
                        foreach (var (xStart, xEnd) in segs)
                        {
                            var bars = SplitSegment(xStart, xEnd, rowIdx2, lap, isTop,
                                piles2, isHoriz: true, perpCoord: y2,
                                bottomLaps: isTop ? bottomLapsX : null);
                            foreach (var (xa, xb) in bars)
                            { DrawBar(tr, btr, xa, y2, xb, y2); result.BarsDrawn++; }
                            for (int i = 0; i < bars.Count - 1; i++)
                                lapPosX.Add((bars[i].b + bars[i + 1].a) / 2.0);
                        }
                        y2 += spacing; rowIdx2++;
                    }

                    double x2 = xMin2 + cover;
                    rowIdx2 = 0;
                    while (x2 <= xMax2 - cover)
                    {
                        var segs = IntersectVertical(vertices, x2);
                        foreach (var (yStart, yEnd) in segs)
                        {
                            var bars = SplitSegment(yStart, yEnd, rowIdx2, lap, isTop,
                                piles2, isHoriz: false, perpCoord: x2,
                                bottomLaps: isTop ? bottomLapsY : null);
                            foreach (var (y1, y2b) in bars)
                            { DrawBar(tr, btr, x2, y1, x2, y2b); result.BarsDrawn++; }
                            for (int i = 0; i < bars.Count - 1; i++)
                                lapPosY.Add((bars[i].b + bars[i + 1].a) / 2.0);
                        }
                        x2 += spacing; rowIdx2++;
                    }

                    tr.Commit();
                    result.LapPositionsX = lapPosX;
                    result.LapPositionsY = lapPosY;
                    result.Log = $"Wygenerowano {result.BarsDrawn} prętów ({(isTop ? "T1/T2" : "B1/B2")}).";
                }
                catch (Exception ex) { result.Error = ex.Message; tr.Abort(); }
            }
            return result;
        }

        // ── Intersection helpers ──────────────────────────────────────────────────

        /// <summary>Returns sorted (xStart, xEnd) pairs where y=y0 crosses the polygon.</summary>
        private static List<(double a, double b)> IntersectHorizontal(
            List<Point3d> vertices, double y0)
        {
            var xs = new List<double>();
            int n = vertices.Count;
            for (int i = 0; i < n; i++)
            {
                var v0 = vertices[i];
                var v1 = vertices[(i + 1) % n];
                if ((v0.Y < y0) != (v1.Y < y0))
                {
                    double t = (y0 - v0.Y) / (v1.Y - v0.Y);
                    xs.Add(v0.X + t * (v1.X - v0.X));
                }
            }
            xs.Sort();
            var result = new List<(double, double)>();
            for (int i = 0; i + 1 < xs.Count; i += 2)
                result.Add((xs[i], xs[i + 1]));
            return result;
        }

        /// <summary>Returns sorted (yStart, yEnd) pairs where x=x0 crosses the polygon.</summary>
        private static List<(double a, double b)> IntersectVertical(
            List<Point3d> vertices, double x0)
        {
            var ys = new List<double>();
            int n = vertices.Count;
            for (int i = 0; i < n; i++)
            {
                var v0 = vertices[i];
                var v1 = vertices[(i + 1) % n];
                if ((v0.X < x0) != (v1.X < x0))
                {
                    double t = (x0 - v0.X) / (v1.X - v0.X);
                    ys.Add(v0.Y + t * (v1.Y - v0.Y));
                }
            }
            ys.Sort();
            var result = new List<(double, double)>();
            for (int i = 0; i + 1 < ys.Count; i += 2)
                result.Add((ys[i], ys[i + 1]));
            return result;
        }

        // ── Bar splitting ─────────────────────────────────────────────────────────

        private static List<(double a, double b)> SplitSegment(
            double start, double end, int rowIdx, double lap, bool isTop,
            List<(Point3d Center, double Radius)> piles,
            bool isHoriz, double perpCoord,
            List<double> bottomLaps)
        {
            double L = end - start;
            if (L <= MaxBarLen)
                return new List<(double, double)> { (start, end) };

            // Stagger offset: 6 positions × 250 mm = 1500 mm cycle
            double lapShift = (rowIdx % 6) * StepLen;
            if (isTop) lapShift += 3 * StepLen; // TOP shifted +750 mm vs BOTTOM

            double midX = start + MaxBarLen / 2.0 + (lapShift % MaxBarLen);

            double minMid = start + lap / 2.0 + StepLen;
            double maxMid = end   - lap / 2.0 - StepLen;
            if (minMid > maxMid) { minMid = start + lap / 2.0; maxMid = end - lap / 2.0; }

            midX = Math.Max(minMid, Math.Min(maxMid, midX));

            // TOP: avoid piles and BOTTOM lap positions
            if (isTop)
            {
                for (int attempt = 0; attempt < 5; attempt++)
                {
                    double px = isHoriz ? midX : perpCoord;
                    double py = isHoriz ? perpCoord : midX;
                    bool overPile    = IsNearPile(px, py, piles);
                    bool onBottomLap = bottomLaps != null &&
                        bottomLaps.Any(bl => Math.Abs(bl - midX) < StepLen);
                    if (!overPile && !onBottomLap) break;
                    midX += StepLen;
                    if (midX > maxMid) { midX = minMid; break; }
                }
            }

            double bar1Len   = RoundUp250(midX + lap / 2.0 - start);
            double bar2Len   = RoundUp250(end  - (midX - lap / 2.0));
            double bar1End   = start + bar1Len;
            double bar2Start = end   - bar2Len;

            var result = new List<(double, double)>();

            if (bar1End - start > MaxBarLen)
                result.AddRange(SplitSegment(start, bar1End, rowIdx, lap, isTop,
                    piles, isHoriz, perpCoord, bottomLaps));
            else
                result.Add((start, bar1End));

            if (end - bar2Start > MaxBarLen)
                result.AddRange(SplitSegment(bar2Start, end, rowIdx + 1, lap, isTop,
                    piles, isHoriz, perpCoord, bottomLaps));
            else
                result.Add((bar2Start, end));

            return result;
        }

        private static double RoundUp250(double length)
            => Math.Ceiling(length / StepLen) * StepLen;

        private static bool IsNearPile(double px, double py,
            List<(Point3d Center, double Radius)> piles)
        {
            var pt = new Point3d(px, py, 0);
            return piles.Any(p => p.Center.DistanceTo(pt) < p.Radius + StepLen);
        }

        // ── Template lookup + COPY command helpers ────────────────────────────────

        /// <summary>
        /// Find the template bar for a given bar length (ceiling to nearest 250 mm).
        /// Falls back to the nearest available length in the catalog.
        /// </summary>
        private static bool TryGetTemplate(Dictionary<int, Point3d> templateBars,
            double barLen, out int foundKey, out Point3d foundPt)
        {
            int targetKey = (int)(Math.Ceiling(barLen / 250.0) * 250);
            if (targetKey < 1250) targetKey = 1250;
            if (targetKey > 6000) targetKey = 6000;

            if (templateBars.TryGetValue(targetKey, out foundPt))
            { foundKey = targetKey; return true; }

            // Find nearest available key in catalog
            foundKey = -1;
            int minDist = int.MaxValue;
            foreach (var k in templateBars.Keys)
            {
                int d = Math.Abs(k - targetKey);
                if (d < minDist) { minDist = d; foundKey = k; }
            }
            if (foundKey > 0) { foundPt = templateBars[foundKey]; return true; }

            foundPt = default(Point3d);
            return false;
        }

        /// <summary>
        /// Appends a DISTRIBUTION command for one bar.
        /// Template bar: left endpoint at <paramref name="tb"/>, length <paramref name="lenKey"/> mm.
        /// Distribution range: from (startX, startY) to (endX, endY).
        /// For B1/T1 (horizontal): startX≠endX, startY=endY=y.
        /// For B2/T2 (vertical):   startX=endX=x, startY≠endY.
        /// ASD orients the bar along the range direction automatically.
        /// Two modal dialogs are handled by AsdDialogAutoCloser (must be started before).
        /// </summary>
        private static void AppendDistribution(System.Text.StringBuilder sb, Point3d tb, int lenKey,
            double startX, double startY, double endX, double endY)
        {
            // Regular window (left→right) to select the template bar
            string w1x = (tb.X - 50.0).ToString("F2", _inv);
            string w1y = (tb.Y - 150.0).ToString("F2", _inv);
            string w2x = (tb.X + lenKey + 50.0).ToString("F2", _inv);
            string w2y = (tb.Y + 150.0).ToString("F2", _inv);
            string sX  = startX.ToString("F2", _inv);
            string sY  = startY.ToString("F2", _inv);
            string eX  = endX.ToString("F2", _inv);
            string eY  = endY.ToString("F2", _inv);

            sb.Append("DISTRIBUTION\n");
            sb.Append(w1x).Append(',').Append(w1y).Append('\n'); // window corner 1
            sb.Append(w2x).Append(',').Append(w2y).Append('\n'); // window corner 2 ("1 found")
            sb.Append("\n");                                       // end "Select objects:"
            // ── Dialog "Reinforcement detailing" appears here → auto-closed with OK ──
            sb.Append(sX).Append(',').Append(sY).Append('\n');   // distribution start point
            sb.Append(eX).Append(',').Append(eY).Append('\n');   // distribution end point
            sb.Append("40\n");                                    // cover (mm)
            sb.Append("200\n");                                   // spacing (mm) → description dialog appears
            // ── Dialog "Reinforcement description" appears here → auto-closed with IDOK ──
        }

        // ── AutoCAD draw helpers ──────────────────────────────────────────────────

        private static void DrawBar(Transaction tr, BlockTableRecord btr,
            double x1, double y1, double x2, double y2)
        {
            var line = new Line(new Point3d(x1, y1, 0), new Point3d(x2, y2, 0));
            line.Layer      = BarLayer;
            line.ColorIndex = 5; // blue ACI 5
            btr.AppendEntity(line);
            tr.AddNewlyCreatedDBObject(line, true);
        }

        private static List<(Point3d Center, double Radius)> FindPilesInSlab(
            Transaction tr, BlockTableRecord btr,
            double xMin, double yMin, double xMax, double yMax,
            List<Point3d> slabVertices)
        {
            var result = new List<(Point3d, double)>();
            foreach (ObjectId id in btr)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead);
                    if (ent is Circle c &&
                        string.Equals(c.Layer, PileLayer, StringComparison.OrdinalIgnoreCase))
                    {
                        var ctr = c.Center;
                        if (ctr.X >= xMin && ctr.X <= xMax &&
                            ctr.Y >= yMin && ctr.Y <= yMax &&
                            PointInPolygon(ctr.X, ctr.Y, slabVertices))
                            result.Add((ctr, c.Radius));
                    }
                }
                catch { }
            }
            return result;
        }

        private static bool PointInPolygon(double px, double py, List<Point3d> verts)
        {
            bool inside = false;
            int n = verts.Count;
            for (int i = 0, j = n - 1; i < n; j = i++)
            {
                double xi = verts[i].X, yi = verts[i].Y;
                double xj = verts[j].X, yj = verts[j].Y;
                if ((yi > py) != (yj > py) &&
                    px < (xj - xi) * (py - yi) / (yj - yi) + xi)
                    inside = !inside;
            }
            return inside;
        }

        private static List<Point3d> GetPolylineVertices(Polyline slab)
        {
            var pts = new List<Point3d>();
            for (int i = 0; i < slab.NumberOfVertices; i++)
                pts.Add(slab.GetPoint3dAt(i));
            return pts;
        }

        private static void GetSlabBounds(List<Point3d> vertices,
            out double xMin, out double yMin, out double xMax, out double yMax)
        {
            xMin = vertices.Min(v => v.X);
            yMin = vertices.Min(v => v.Y);
            xMax = vertices.Max(v => v.X);
            yMax = vertices.Max(v => v.Y);
        }

        private static void EnsureLayer(Transaction tr, Database db,
            string name, short colorIndex)
        {
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(name))
            {
                var ltr = new LayerTableRecord();
                ltr.Name  = name;
                ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
                lt.UpgradeOpen();
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }
    }
}
