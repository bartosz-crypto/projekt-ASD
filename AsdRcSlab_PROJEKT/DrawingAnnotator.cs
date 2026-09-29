using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AsdRcSlab
{
    /// <summary>
    /// Annotates pile circles on the active AutoCAD drawing with PH labels and SOLID hatch fills.
    /// - NO ACTION piles are skipped (no annotation added)
    /// - Copies text style / layer from any existing PH annotations in the drawing
    /// - Falls back to layer "AP rebar top" / style "WYG_0MS" if nothing found
    /// - PH text is placed centred INSIDE the pile circle
    /// - SOLID hatch fills the pile circle; colour depends on PH level
    /// </summary>
    public static class DrawingAnnotator
    {
        public const string LayerPhText  = "AP rebar top";
        public const string LayerPhHatch = "AP-Hatch";
        public const string LayerNotUsed = "AP-NOTUSED";

        // Default text style name used in Speedeck drawings
        private const string DefaultTextStyle = "WYG_0MS";

        public static AnnotationResult Annotate(List<PileData> piles)
        {
            var result = new AnnotationResult();

            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) { result.Log = "Brak aktywnego dokumentu."; return result; }

            var db = doc.Database;

            // ── Pre-check: ensure this is a Speedeck drawing before modifying anything
            if (!CheckSpeeddeckTitle(db))
            {
                result.WrongDrawing = true;
                result.Log = "Drawing does not contain the title 'SPEEDECK PILED RAFT FOUNDATION'.\n" +
                             "Make sure the active document is the correct reinforcement drawing.";
                return result;
            }

            // ── Cleanup encji z poprzedniego runu zanim dodamy nowe (idempotent)
            CleanupPreviousAnnotations(db);

            var manualPileIds = new List<string>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt  = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // ── 1. Collect drawing entities ──────────────────────────────────────
                var allTexts   = new List<(string Text, Point3d Pos, ObjectId Id, double Height, string Style, string Layer)>();
                var allCircles = new List<(Point3d Center, double Radius, ObjectId Id, string Layer)>();
                // p163: pale KWADRATOWE 200x200 (LWPOLYLINE zamknięta) — hatchowane jak okręgi.
                var allSquares = new List<(Point3d Center, double HalfSide, ObjectId Id, string Layer)>();

                foreach (ObjectId id in btr)
                {
                    try
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead);

                        if (ent is DBText dbt)
                        {
                            allTexts.Add((dbt.TextString?.Trim() ?? "",
                                          dbt.Position, id,
                                          dbt.Height,
                                          dbt.TextStyleName ?? "",
                                          dbt.Layer ?? ""));
                        }
                        else if (ent is MText mt)
                        {
                            string plain = StripMTextFormat(mt.Contents ?? "");
                            allTexts.Add((plain, mt.Location, id,
                                          mt.TextHeight,
                                          mt.TextStyleName ?? "",
                                          mt.Layer ?? ""));
                        }
                        else if (ent is Circle c)
                        {
                            allCircles.Add((c.Center, c.Radius, id, c.Layer ?? ""));
                        }
                        else if (ent is Polyline pl && TryGetSquarePile(pl, out var sqCenter, out var sqHalf))
                        {
                            allSquares.Add((sqCenter, sqHalf, id, pl.Layer ?? ""));
                        }
                    }
                    catch { /* proxy or locked entities */ }
                }

                result.TotalCircles = allCircles.Count;
                result.TotalSquares = allSquares.Count;
                result.TotalTexts   = allTexts.Count;
                int squaresHatched  = 0;

                // ── 2. Detect existing PH annotation style ───────────────────────────
                string detectedLayer = LayerPhText;
                string detectedStyle = DefaultTextStyle;
                double detectedHeightFactor = 0.8; // text height = radius × factor

                var existingPh = allTexts
                    .Where(t => Regex.IsMatch(t.Text, @"^PH\d", RegexOptions.IgnoreCase))
                    .ToList();

                if (existingPh.Count > 0)
                {
                    var sample = existingPh[0];
                    if (!string.IsNullOrEmpty(sample.Layer)) detectedLayer = sample.Layer;
                    if (!string.IsNullOrEmpty(sample.Style)) detectedStyle = sample.Style;
                    // Find nearest circle to estimate height factor
                    var nearCircle = allCircles
                        .OrderBy(c => c.Center.DistanceTo(sample.Pos))
                        .FirstOrDefault();
                    if (nearCircle.Radius > 0 && sample.Height > 0)
                        detectedHeightFactor = sample.Height / nearCircle.Radius;
                    result.Log += $"Existing PH style detected: layer={detectedLayer}, style={detectedStyle}\n";
                }

                // Ensure annotation layers and text style exist
                EnsureLayer(tr, db, detectedLayer, 2);   // yellow
                EnsureLayer(tr, db, LayerPhHatch, 256);  // byLayer
                string styleId = EnsureTextStyle(tr, db, detectedStyle);

                btr.UpgradeOpen();

                // ── 3. Annotate each pile ─────────────────────────────────────────────
                foreach (var pile in piles)
                {
                    // MANUAL → collect for jig insertion, no drawing annotation
                    if (string.Equals(pile.PhAction, "MANUAL", StringComparison.OrdinalIgnoreCase))
                    {
                        manualPileIds.Add(pile.PileId);
                        continue;
                    }

                    // NO ACTION → no annotation
                    if (pile.PhAction == "NO ACTION" ||
                        string.IsNullOrEmpty(pile.PhAction))
                    {
                        result.Skipped.Add(pile.PileId);
                        continue;
                    }

                    // Find text entity matching this pile ID
                    // p165: predykat "czy obok tekstu jest pal" — FindPileText preferuje
                    // etykiety stojące przy okręgu/kwadracie pala, a nie np. znaczniki
                    // prętów "103"/"107"/"108" z warstwy 0-25TEXT.
                    var matchText = FindPileText(pile.PileId, allTexts,
                        pos => allCircles.Any(c => c.Center.DistanceTo(pos) < c.Radius * 8) ||
                               allSquares.Any(s => s.Center.DistanceTo(pos) < s.HalfSide * 8));
                    if (matchText == null)
                    {
                        result.NotFound.Add(pile.PileId);
                        continue;
                    }

                    Point3d textPos = matchText.Value.Pos;

                    // Find nearest circle to pile text
                    var nearCircle = allCircles
                        .Where(c => c.Center.DistanceTo(textPos) < c.Radius * 8)
                        .OrderBy(c => c.Center.DistanceTo(textPos))
                        .FirstOrDefault();

                    double radius;
                    Point3d center;

                    if (nearCircle.Id != ObjectId.Null)
                    {
                        radius = nearCircle.Radius;
                        center = nearCircle.Center;

                        // ── Add SOLID hatch inside circle ─────────────────────────
                        AddSolidHatch(tr, db, btr, center, radius, pile.PhAction);
                    }
                    else
                    {
                        // p163: brak okręgu → spróbuj pal KWADRATOWY (LWPOLYLINE 200x200)
                        var nearSquare = allSquares
                            .Where(s => s.Center.DistanceTo(textPos) < s.HalfSide * 8)
                            .OrderBy(s => s.Center.DistanceTo(textPos))
                            .FirstOrDefault();

                        if (nearSquare.Id == ObjectId.Null)
                        {
                            result.NotFound.Add($"{pile.PileId}(no circle)");
                            continue;
                        }

                        radius = nearSquare.HalfSide;       // pseudo-promień do tekstu
                        center = nearSquare.Center;

                        // ── Hatch kwadratu — REUŻYWA rdzenia hatcha (1:1 z okręgiem),
                        //    granica = istniejący LWPOLYLINE pala ───────────────────
                        AddSolidHatchOnBoundary(tr, btr,
                            new ObjectIdCollection { nearSquare.Id }, pile.PhAction);
                        squaresHatched++;
                    }

                    // ── Add PH text centred inside circle ─────────────────────────
                    double textHeight = radius * detectedHeightFactor;
                    if (textHeight < 50)  textHeight = 50;
                    if (textHeight > 300) textHeight = 300;

                    // MText format matching Speedeck style: PH{font;number}
                    string mtContent = FormatPhMText(pile.PhAction);

                    var mt2 = new MText();
                    mt2.SetDatabaseDefaults();
                    mt2.TextHeight     = textHeight;
                    mt2.TextStyleId    = GetTextStyleId(tr, db, detectedStyle);
                    mt2.Attachment     = AttachmentPoint.TopLeft;
                    // Place PH text directly below the pile ID text, left-aligned with it
                    mt2.Location       = new Point3d(textPos.X, textPos.Y - textHeight * 0.2, 0);
                    mt2.Width          = radius * 3.0;
                    mt2.Contents       = mtContent;
                    mt2.Layer          = detectedLayer;
                    mt2.Color          = Color.FromColorIndex(ColorMethod.ByAci, 4); // cyan
                    btr.AppendEntity(mt2);
                    tr.AddNewlyCreatedDBObject(mt2, true);

                    result.Annotated.Add(pile.PileId);
                }

                var phLogSb = new StringBuilder();
                var (phLabels, unusedMarked) = AnnotatePhDetailLabels(tr, btr, db, phLogSb);
                result.PhLabelsUpdated = phLabels;
                result.UnusedMarked    = unusedMarked;
                result.Log += phLogSb.ToString();

                result.SquaresHatched = squaresHatched;

                tr.Commit();
            }

            // p163: stamp widoczny w konsoli
            doc.Editor.WriteMessage($"\n[PAA] squares hatched: {result.SquaresHatched}\n");

            result.ManualPileIds = manualPileIds;
            result.Log += BuildLog(result);
            return result;
        }

        // ── Cleanup helpers ──────────────────────────────────────────────────────────

        private static bool CheckSpeeddeckTitle(Database db)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt  = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                return HasSpeeddeckTitle(tr, btr);
                // Dispose without Commit = Abort; fine for read-only
            }
        }

        private static void CleanupPreviousAnnotations(Database db)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var toDelete = new List<ObjectId>();
                foreach (ObjectId id in ms)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;
                    if (string.Equals(ent.Layer, LayerPhText,  StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(ent.Layer, LayerPhHatch, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(ent.Layer, LayerNotUsed, StringComparison.OrdinalIgnoreCase))
                    {
                        toDelete.Add(id);
                    }
                }

                foreach (var id in toDelete)
                {
                    var ent = (Entity)tr.GetObject(id, OpenMode.ForWrite);
                    ent.Erase();
                }

                tr.Commit();
            }
        }

        // ── AP-TEXT template updater ─────────────────────────────────────────────────

        private static (int PhLabelsUpdated, int UnusedMarked) AnnotatePhDetailLabels(
            Transaction tr, BlockTableRecord btr, Database db, StringBuilder log)
        {
            var phRegex   = new Regex(@"\bPH([1-9])\b");
            // Loosened: matches anything from '(' through 'No...LOCATIONS?' to ')'.
            // Original strict regex failed when MText encoding inserted invisible chars
            // between the digit and 'No' (e.g. formatting codes or non-breaking space).
            var locRegex  = new Regex(@"\(\s*\d*[^)]*No[^)]*LOCATIONS?[^)]*\)", RegexOptions.IgnoreCase);
            var applRegex = new Regex(@"APPLICABLE FOR PILES?[^}]*");

            int updatedCount = 0;
            int unusedMarked = 0;
            int skippedNoPh  = 0;

            foreach (ObjectId id in btr)
            {
                // Open ForWrite directly — avoids UpgradeOpen() edge cases on iterated BTR entities
                MText ent;
                try { ent = tr.GetObject(id, OpenMode.ForWrite) as MText; }
                catch { continue; }
                if (ent == null) continue;
                if (!string.Equals(ent.Layer, "AP-TEXT", StringComparison.OrdinalIgnoreCase)) continue;

                string contents = ent.Contents;
                var phMatch = phRegex.Match(contents);
                if (!phMatch.Success) { skippedNoPh++; continue; }

                string phKey = "PH" + phMatch.Groups[1].Value;

                var piles = SessionData.Piles
                    .Where(p => string.Equals(p.PhAction, phKey, StringComparison.OrdinalIgnoreCase))
                    .Select(p => p.PileId)
                    .OrderBy(pid => ExtractFirstNumber(pid))
                    .ThenBy(pid => pid, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // Zawsze updateuj — również gdy brak pali (0No LOCATIONS / APPLICABLE FOR PILE/S —)
                string locReplacement = piles.Count == 1
                    ? "(1No LOCATION)"
                    : $"({piles.Count}No LOCATIONS)";
                contents = locRegex.Replace(contents, locReplacement);

                string pileListJoined = string.Join(", ", piles);
                string applReplacement;
                if (piles.Count == 0)
                    applReplacement = "APPLICABLE FOR PILE/S —";
                else if (piles.Count == 1)
                    applReplacement = $"APPLICABLE FOR PILE {pileListJoined}";
                else
                    applReplacement = $"APPLICABLE FOR PILES {pileListJoined}";
                contents = applRegex.Replace(contents, applReplacement);

                ent.Contents = contents;

                string logSuffix = piles.Count > 0 ? $": {pileListJoined}" : "";
                log.AppendLine($"  AP-TEXT [{phKey}]: zaktualizowano ({piles.Count} pali{logSuffix})");
                updatedCount++;

                // Gdy PH bez pali — narysuj czerwony krzyż przez bbox szablonu
                if (piles.Count == 0)
                {
                    try
                    {
                        EnsureLayer(tr, db, LayerNotUsed, 1);
                        var ext = ent.GeometricExtents;
                        double x1 = ext.MinPoint.X, y1 = ext.MinPoint.Y;
                        double x2 = ext.MaxPoint.X, y2 = ext.MaxPoint.Y;

                        var line1 = new Line(new Point3d(x1, y1, 0), new Point3d(x2, y2, 0));
                        line1.Layer      = LayerNotUsed;
                        line1.ColorIndex = 1;
                        btr.AppendEntity(line1);
                        tr.AddNewlyCreatedDBObject(line1, true);

                        var line2 = new Line(new Point3d(x1, y2, 0), new Point3d(x2, y1, 0));
                        line2.Layer      = LayerNotUsed;
                        line2.ColorIndex = 1;
                        btr.AppendEntity(line2);
                        tr.AddNewlyCreatedDBObject(line2, true);

                        unusedMarked++;
                    }
                    catch (System.Exception ex)
                    {
                        AcApp.DocumentManager.MdiActiveDocument?.Editor
                            ?.WriteMessage($"\nAP-NOTUSED cross fail for {phKey}: {ex.Message}");
                    }
                }
            }

            log.AppendLine($"AnnotatePhDetailLabels: updated {updatedCount}, skipped " +
                           $"{skippedNoPh} (no PH in content).");
            return (updatedCount, unusedMarked);
        }

        private static int ExtractFirstNumber(string s)
        {
            if (string.IsNullOrEmpty(s)) return int.MaxValue;
            var m = Regex.Match(s, @"\d+");
            if (!m.Success) return int.MaxValue;
            return int.TryParse(m.Value, out int n) ? n : int.MaxValue;
        }

        // ── Private helpers ──────────────────────────────────────────────────────────

        private static bool HasSpeeddeckTitle(Transaction tr, BlockTableRecord btr)
        {
            // Search raw MText/DBText content for the reinforcement drawing title.
            // "REINFORCEMENT DETAILS OF SPEEDECK" appears on one line (before any \P break)
            // so it is always a literal substring in the raw content — no stripping needed.
            // This phrase is specific to the detail sheet, not the PLOT layout drawing.
            const string key = "REINFORCEMENT DETAILS OF SPEEDECK";
            foreach (ObjectId id in btr)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead);
                    string raw = null;
                    if (ent is DBText dbt) raw = dbt.TextString;
                    else if (ent is MText mt) raw = mt.Contents;
                    if (!string.IsNullOrEmpty(raw) && raw.ToUpperInvariant().Contains(key))
                        return true;
                }
                catch { }
            }
            return false;
        }

        private static void AddSolidHatch(Transaction tr, Database db,
            BlockTableRecord btr, Point3d center, double radius, string phAction)
        {
            // Boundary circle (invisible) — własna granica na warstwie hatcha (kasowana przy cleanup)
            var boundCircle = new Circle(center, Vector3d.ZAxis, radius);
            boundCircle.SetDatabaseDefaults();
            boundCircle.Layer   = LayerPhHatch;
            boundCircle.Visible = false;
            btr.AppendEntity(boundCircle);
            tr.AddNewlyCreatedDBObject(boundCircle, true);

            AddSolidHatchOnBoundary(tr, btr,
                new ObjectIdCollection { boundCircle.ObjectId }, phAction);
        }

        // p163: rdzeń tworzenia hatcha — JEDYNE miejsce z parametrami wzoru (pattern/scale/
        // angle/color/warstwa/associative). Granica przekazana jako argument → ten sam styl
        // 1:1 dla okręgów i kwadratów. Różni się TYLKO pętla granicy.
        private static void AddSolidHatchOnBoundary(Transaction tr,
            BlockTableRecord btr, ObjectIdCollection loop, string phAction)
        {
            var hatch = new Hatch();
            hatch.SetDatabaseDefaults();
            hatch.HatchObjectType = HatchObjectType.HatchObject;
            hatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI31");
            hatch.PatternScale = 10.0;
            hatch.ColorIndex   = 1;   // red
            hatch.Layer        = LayerPhHatch;
            btr.AppendEntity(hatch);
            tr.AddNewlyCreatedDBObject(hatch, true);
            hatch.Associative = false;
            hatch.AppendLoop(HatchLoopTypes.Outermost, loop);
            hatch.EvaluateHatch(true);
        }

        // p163: kryterium pala KWADRATOWEGO — LWPOLYLINE zamknięta, 4 wierzchołki, proste
        // segmenty (brak bulge), bbox ~200x200 (tolerancja 180..220 na bok). Zwraca środek
        // bbox i pół-bok (pseudo-promień do umieszczenia tekstu PH).
        private static bool TryGetSquarePile(Polyline pl, out Point3d center, out double halfSide)
        {
            center   = Point3d.Origin;
            halfSide = 0;
            if (pl == null || !pl.Closed || pl.NumberOfVertices != 4) return false;

            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            for (int i = 0; i < pl.NumberOfVertices; i++)
            {
                if (Math.Abs(pl.GetBulgeAt(i)) > 1e-6) return false;   // łuk → nie kwadrat
                var p = pl.GetPoint2dAt(i);
                if (p.X < minX) minX = p.X;
                if (p.X > maxX) maxX = p.X;
                if (p.Y < minY) minY = p.Y;
                if (p.Y > maxY) maxY = p.Y;
            }

            double w = maxX - minX, h = maxY - minY;
            if (w < 180 || w > 220 || h < 180 || h > 220) return false;

            center   = new Point3d((minX + maxX) / 2.0, (minY + maxY) / 2.0, 0);
            halfSide = Math.Max(w, h) / 2.0;
            return true;
        }

        private static string EnsureTextStyle(Transaction tr, Database db, string styleName)
        {
            var tt = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
            if (!tt.Has(styleName))
            {
                // Create a minimal style (uses "romans.shx" — common in AutoCAD)
                var ttr = new TextStyleTableRecord();
                ttr.Name     = styleName;
                ttr.FileName = "romans.shx";
                tt.UpgradeOpen();
                tt.Add(ttr);
                tr.AddNewlyCreatedDBObject(ttr, true);
            }
            return styleName;
        }

        private static ObjectId GetTextStyleId(Transaction tr, Database db, string styleName)
        {
            var tt = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
            if (tt.Has(styleName)) return tt[styleName];
            return db.Textstyle; // current style as fallback
        }

        private static void EnsureLayer(Transaction tr, Database db, string name, short colorIndex)
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

        private static (string Text, Point3d Pos, ObjectId Id, double Height, string Style, string Layer)?
            FindPileText(string pileId,
                         List<(string Text, Point3d Pos, ObjectId Id, double Height, string Style, string Layer)> texts,
                         Func<Point3d, bool> hasPileNearby)
        {
            // p165 BUG: stara wersja brała PIERWSZY tekst po NormalizePileId (obcina 'P'),
            // więc "P103" łapało znacznik pręta "103" (warstwa 0-25TEXT) z detali zbrojenia,
            // obok którego nie ma okręgu pala → "Not found (no circle)" i pal niepodpisany.
            // Kolejność kandydatów:
            //   1. dokładne dopasowanie ("P103" == "P103", case-insensitive) przy palu
            //   2. dopasowanie znormalizowane ("103" ~ "P103") przy palu
            //   3. dokładne dopasowanie gdziekolwiek
            //   4. dopasowanie znormalizowane gdziekolwiek (legacy)
            string id     = (pileId ?? "").Trim();
            string normId = NormalizePileId(id);

            var exact = texts
                .Where(t => string.Equals(t.Text, id, StringComparison.OrdinalIgnoreCase))
                .ToList();
            var normalized = texts
                .Where(t => NormalizePileId(t.Text) == normId)
                .ToList();

            foreach (var group in new[] { exact, normalized })
            {
                var nearPile = group.Where(t => hasPileNearby(t.Pos)).ToList();
                if (nearPile.Count > 0) return nearPile[0];
            }

            if (exact.Count > 0)      return exact[0];
            if (normalized.Count > 0) return normalized[0];
            return null;
        }

        private static string NormalizePileId(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim().TrimStart('P', 'p');
            if (int.TryParse(s, out int n)) return n.ToString();
            return s.ToUpperInvariant();
        }

        private static string StripMTextFormat(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "";
            string s = raw.Replace("\\P", " ").Replace("\\p", " ");
            s = Regex.Replace(s, @"\\\w[^;]*;", "");
            s = Regex.Replace(s, @"\{[^}]*\}", m =>
            {
                // Keep text content inside braces
                string inner = Regex.Replace(m.Value.Trim('{', '}'), @"\\\w[^;]*;", "");
                return inner;
            });
            return s.Trim();
        }

        /// <summary>Build MText content string matching Speedeck format.</summary>
        private static string FormatPhMText(string phAction)
        {
            // Format: PH + number in Romans font, e.g. "PH{\Fromans|c238;3}"
            var m = Regex.Match(phAction, @"\d+");
            string suffix = m.Success ? m.Value : phAction.Replace("PH", "");
            return $@"PH{{\Fromans|c238;{suffix}}}";
        }

        private static short PhColorIndex(string ph)
        {
            // ACI colour by PH level (matching existing Speedeck colour scheme)
            switch (ph)
            {
                case "PH1": case "PH2": case "PH3": return 3;   // green
                case "PH4": case "PH5": case "PH6": return 2;   // yellow
                case "PH7": case "PH8": case "PH9": return 1;   // red
                default:                            return 7;   // white
            }
        }

        private static string BuildLog(AnnotationResult r)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Labelled: {r.Annotated.Count} | Skipped NO ACTION: {r.Skipped.Count} | Not found: {r.NotFound.Count}");
            sb.AppendLine($"Squares hatched: {r.SquaresHatched} (of {r.TotalSquares} detected)");
            if (r.NotFound.Count > 0)
                sb.AppendLine($"Not found: {string.Join(", ", r.NotFound)}");
            if (r.UnusedMarked > 0)
                sb.AppendLine($"Unused details (cross): {r.UnusedMarked}");
            return sb.ToString();
        }
    }

    public class AnnotationResult
    {
        public int          TotalCircles     { get; set; }
        public int          TotalSquares     { get; set; }
        public int          SquaresHatched   { get; set; }
        public int          TotalTexts       { get; set; }
        public List<string> Annotated        { get; set; } = new List<string>();
        public List<string> Skipped          { get; set; } = new List<string>();
        public List<string> NotFound         { get; set; } = new List<string>();
        public List<string> ManualPileIds    { get; set; } = new List<string>();
        public string       Log              { get; set; } = "";
        public bool         WrongDrawing     { get; set; }
        public int          PhLabelsUpdated  { get; set; }
        public int          UnusedMarked     { get; set; }
    }
}
