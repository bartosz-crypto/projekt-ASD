using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
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
                    tr.Commit();
                }
                catch (Exception ex)
                {
                    result.Error = ex.Message;
                    tr.Abort();
                    return result;
                }
            }

            // ── 2. Build DISTRIBUTION command string ─────────────────────────────
            // Zone-based approach: scan slab at spacing intervals and group consecutive
            // scanlines with the same intersection extents into zones. Each zone gets one
            // DISTRIBUTION call with the correct template bar length.
            //   B1/T1: scan horizontal lines (Y axis), zone = strip of equal width bars
            //   B2/T2: scan vertical lines (X axis), zone = strip of equal height bars

            if (templateBars.Count == 0)
            {
                result.Error = "Brak szablonów prętów. Uruchom najpierw ASD-GSETUP.";
                return result;
            }

            var cmdSb    = new System.Text.StringBuilder();
            int callCount = 0;
            var dbg = new System.Text.StringBuilder();

            // B1/T1: horizontal bars, distribution axis = vertical
            var hZones = DetectZones(vertices, yMin + cover, yMax - cover, spacing, isHorizontal: true);
            dbg.AppendLine($"B1 zones ({hZones.Count}):");
            foreach (var (xS, xE, yS, yE) in hZones)
            {
                double barLen = xE - xS;
                dbg.AppendLine($"  xS={xS:F0} xE={xE:F0} barLen={barLen:F0}  yS={yS:F0} yE={yE:F0}");
                int n = AppendZoneDistributions(cmdSb, templateBars,
                    xS, xE, yS, yE, isHorizontal: true, lap);
                callCount += n;
                dbg.AppendLine($"    → {n} call(s)");
            }

            // B2/T2: vertical bars, distribution axis = horizontal
            var vZones = DetectZones(vertices, xMin + cover, xMax - cover, spacing, isHorizontal: false);
            dbg.AppendLine($"B2 zones ({vZones.Count}):");
            foreach (var (yS, yE, xS, xE) in vZones)
            {
                double barLen = yE - yS;
                dbg.AppendLine($"  yS={yS:F0} yE={yE:F0} barLen={barLen:F0}  xS={xS:F0} xE={xE:F0}");
                int n = AppendZoneDistributions(cmdSb, templateBars,
                    yS, yE, xS, xE, isHorizontal: false, lap);
                callCount += n;
                dbg.AppendLine($"    → {n} call(s)");
            }

            var lapPosX = new List<double>();
            var lapPosY = new List<double>();

            // ── Debug dump ────────────────────────────────────────────────────────
            try
            {
                var header = new System.Text.StringBuilder();
                header.AppendLine($"vertices: {vertices.Count}");
                foreach (var v in vertices)
                    header.AppendLine($"  ({v.X:F1}, {v.Y:F1})");
                header.AppendLine($"bounds: x=[{xMin:F1},{xMax:F1}]  y=[{yMin:F1},{yMax:F1}]");
                header.AppendLine($"cover={cover}  spacing={spacing}  callCount={callCount}");
                header.Append(dbg);
                header.AppendLine("--- command string ---");
                string cs = cmdSb.ToString();
                header.Append(cs.Length > 3000 ? cs.Substring(0, 3000) : cs);
                System.IO.File.WriteAllText(
                    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "asd_gbot_debug.txt"),
                    header.ToString());
            }
            catch { }

            // ── 3. Send commands line-by-line with delays from a background thread ──
            if (callCount > 0)
            {
                var lines = cmdSb.ToString().Split('\n');
                doc.Editor.WriteMessage($"\nGBOT: {callCount} stref DISTRIBUTION ({(isTop ? "T1/T2" : "B1/B2")})...");

                AsdDialogAutoCloser.Start(timeoutMs: 1_200_000);

                System.Threading.Tasks.Task.Run(() =>
                {
                    for (int li = 0; li < lines.Length; li++)
                    {
                        if (li == lines.Length - 1 && lines[li] == "") continue;

                        doc.SendStringToExecute(lines[li] + "\n", true, false, false);

                        var trimmed = lines[li].Trim();
                        int delayMs;
                        bool nextIsDistrib = li + 1 < lines.Length &&
                            lines[li + 1].Trim().Equals("DISTRIBUTION",
                                System.StringComparison.OrdinalIgnoreCase);
                        if (trimmed.Equals("DISTRIBUTION", System.StringComparison.OrdinalIgnoreCase))
                            delayMs = 300;
                        else if (trimmed == "")
                            delayMs = 500;  // blank ends selection → dialog 1 appears
                        else if (nextIsDistrib || li == lines.Length - 1)
                            delayMs = 600;  // "200" → dialog 2 appears
                        else
                            delayMs = 80;
                        System.Threading.Thread.Sleep(delayMs);
                    }
                    System.Threading.Thread.Sleep(2000);
                    AsdDialogAutoCloser.Stop();
                });
            }

            result.BarsDrawn     = callCount;
            result.LapPositionsX = lapPosX;
            result.LapPositionsY = lapPosY;
            result.Log = $"Wysłano {callCount} komend DISTRIBUTION ({(isTop ? "T1/T2" : "B1/B2")}).";

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

                    var piles2 = new List<(Point3d Center, double Radius)>(); // pile avoidance not used in line-based fallback
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

        // ── Zone detection ────────────────────────────────────────────────────────

        /// <summary>
        /// Scans the slab polygon at <paramref name="scanStep"/> intervals and groups
        /// consecutive scan positions with the same intersection segments into zones.
        ///
        /// isHorizontal=true  → scans horizontal lines (Y varies); returns (xS,xE,yS,yE):
        ///   xS..xE = bar extent (horizontal), yS..yE = distribution axis (vertical).
        /// isHorizontal=false → scans vertical lines (X varies); returns (yS,yE,xS,xE):
        ///   yS..yE = bar extent (vertical), xS..xE = distribution axis (horizontal).
        /// </summary>
        private static List<(double segS, double segE, double axisS, double axisE)> DetectZones(
            List<Point3d> vertices,
            double scanMin, double scanMax, double scanStep,
            bool isHorizontal)
        {
            var result = new List<(double, double, double, double)>();
            List<(double a, double b)> curSegs = null;
            double zoneAxisStart = 0;
            double zoneAxisEnd   = 0;

            for (double scan = scanMin; scan <= scanMax + 1e-6; scan += scanStep)
            {
                var segs = isHorizontal
                    ? IntersectHorizontal(vertices, scan)
                    : IntersectVertical(vertices, scan);

                if (segs.Count == 0) continue;

                if (curSegs == null || !SegmentsMatch(curSegs, segs))
                {
                    // Flush previous zone
                    if (curSegs != null)
                        foreach (var (a, b) in curSegs)
                            result.Add((a, b, zoneAxisStart, zoneAxisEnd));

                    curSegs       = segs;
                    zoneAxisStart = scan;
                    zoneAxisEnd   = scan;
                }
                else
                {
                    zoneAxisEnd = scan;
                }
            }

            // Flush last zone
            if (curSegs != null)
                foreach (var (a, b) in curSegs)
                    result.Add((a, b, zoneAxisStart, zoneAxisEnd));

            return result;
        }

        private static bool SegmentsMatch(
            List<(double a, double b)> s1,
            List<(double a, double b)> s2,
            double tol = 100.0)
        {
            if (s1.Count != s2.Count) return false;
            for (int i = 0; i < s1.Count; i++)
                if (Math.Abs(s1[i].a - s2[i].a) > tol || Math.Abs(s1[i].b - s2[i].b) > tol)
                    return false;
            return true;
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
        /// Appends one or more DISTRIBUTION commands to cover a zone.
        /// For zones wider/taller than MaxBarLen, splits into overlapping sub-distributions
        /// with <paramref name="lap"/> mm overlap (lap joints).
        ///
        /// isHorizontal=true  → B1/T1: segStart/segEnd = X extents of each bar,
        ///                                axisStart/axisEnd = Y range of distribution axis.
        /// isHorizontal=false → B2/T2: segStart/segEnd = Y extents of each bar,
        ///                                axisStart/axisEnd = X range of distribution axis.
        /// Returns number of DISTRIBUTION calls appended.
        /// </summary>
        private static int AppendZoneDistributions(System.Text.StringBuilder sb,
            Dictionary<int, Point3d> templateBars,
            double segStart, double segEnd,
            double axisStart, double axisEnd,
            bool isHorizontal, double lap)
        {
            int count = 0;
            double pos = segStart;

            while (true)
            {
                double remaining = segEnd - pos;
                double barLen = remaining <= MaxBarLen + 1.0
                    ? RoundUp250(remaining)
                    : MaxBarLen;

                double barEnd    = pos + barLen;
                double axisCoord = (pos + barEnd) / 2.0;

                if (TryGetTemplate(templateBars, barLen, out int lk, out Point3d tb))
                {
                    if (isHorizontal)
                        AppendDistribution(sb, tb, lk, axisCoord, axisStart, axisCoord, axisEnd);
                    else
                        AppendDistribution(sb, tb, lk, axisStart, axisCoord, axisEnd, axisCoord);
                    count++;
                }

                if (remaining <= MaxBarLen + 1.0) break;    // last segment done
                pos = pos + MaxBarLen - lap;                 // next segment overlaps by lap
            }

            return count;
        }

        /// <summary>
        /// Appends a DISTRIBUTION command for one bar (B1/B2/T1/T2).
        /// Template bar: left endpoint at <paramref name="tb"/>, length <paramref name="lenKey"/> mm.
        ///
        /// Uses a short 200 mm distribution range with cover=40 and spacing=200:
        ///   available = 200 - 2×40 = 120 mm  &lt;  spacing=200  →  exactly 1 bar per call.
        /// The bar is placed 40 mm from the range start (i.e., at the given start point + 40 mm)
        /// and extends for lenKey mm in the range direction.
        /// This avoids the ASD maximum-spacing rejection that occurs with large spacing values.
        ///
        /// Two modal dialogs are auto-closed by AsdDialogAutoCloser (must be started before).
        /// ASD DISTRIBUTION sequence (no pre-selection — bar selected via window):
        ///   DISTRIBUTION
        ///   Select objects: w1x,w1y  w2x,w2y  [blank]  → "1 found"
        ///   [Dialog "Reinforcement detailing" → auto-closed]
        ///   Start distribution line: startX,startY
        ///   End distribution point:  endX,endY
        ///   Position of first bar (mm): 40
        ///   &lt;number&gt;*&lt;spacing(mm)&gt;: 200
        ///   [Dialog "Reinforcement description" → auto-closed → bars placed]
        /// </summary>
        private static void AppendDistribution(System.Text.StringBuilder sb, Point3d tb, int lenKey,
            double startX, double startY, double endX, double endY)
        {
            // Window that encloses the template bar (left-to-right = regular window)
            string w1x = (tb.X - 50.0).ToString("F2", _inv);
            string w1y = (tb.Y - 30.0).ToString("F2", _inv);
            string w2x = (tb.X + lenKey + 50.0).ToString("F2", _inv);
            string w2y = (tb.Y + 30.0).ToString("F2", _inv);
            string sX  = startX.ToString("F2", _inv);
            string sY  = startY.ToString("F2", _inv);
            string eX  = endX.ToString("F2", _inv);
            string eY  = endY.ToString("F2", _inv);

            sb.Append("DISTRIBUTION\n");
            sb.Append("W\n");                                      // force Window selection mode
            sb.Append(w1x).Append(',').Append(w1y).Append('\n'); // first corner
            sb.Append(w2x).Append(',').Append(w2y).Append('\n'); // second corner → "1 found"
            sb.Append("\n");                                       // end selection
            // [Dialog "Reinforcement detailing" → auto-closed by AsdDialogAutoCloser]
            sb.Append(sX).Append(',').Append(sY).Append('\n');   // Start distribution line
            sb.Append(eX).Append(',').Append(eY).Append('\n');   // End distribution point
            sb.Append("40\n");                                     // Position of first bar (mm)
            sb.Append("200\n");                                    // <number>*<spacing(mm)>
            // [Dialog "Reinforcement description" → auto-closed → bars placed]
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
            Document doc,
            double xMin, double yMin, double xMax, double yMax,
            List<Point3d> slabVertices)
        {
            var result = new List<(Point3d, double)>();
            try
            {
                // Use SelectCrossingWindow with layer filter — uses AutoCAD's spatial index,
                // much faster than iterating all ModelSpace entities.
                var filter = new SelectionFilter(new[]
                {
                    new TypedValue((int)DxfCode.Start,      "CIRCLE"),
                    new TypedValue((int)DxfCode.LayerName,  PileLayer),
                });
                var pt1 = new Point3d(xMin, yMin, 0);
                var pt2 = new Point3d(xMax, yMax, 0);
                var selRes = doc.Editor.SelectCrossingWindow(pt1, pt2, filter);
                if (selRes.Status != PromptStatus.OK) return result;

                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    foreach (var id in selRes.Value.GetObjectIds())
                    {
                        var c = tr.GetObject(id, OpenMode.ForRead) as Circle;
                        if (c != null && PointInPolygon(c.Center.X, c.Center.Y, slabVertices))
                            result.Add((c.Center, c.Radius));
                    }
                    tr.Commit();
                }
            }
            catch { }
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
