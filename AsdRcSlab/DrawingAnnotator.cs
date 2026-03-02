using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
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

        // Default text style name used in Speedeck drawings
        private const string DefaultTextStyle = "WYG_0MS";

        public static AnnotationResult Annotate(List<PileData> piles)
        {
            var result = new AnnotationResult();

            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) { result.Log = "Brak aktywnego dokumentu."; return result; }

            var db = doc.Database;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt  = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // ── 1. Collect drawing entities ──────────────────────────────────────
                var allTexts   = new List<(string Text, Point3d Pos, ObjectId Id, double Height, string Style, string Layer)>();
                var allCircles = new List<(Point3d Center, double Radius, ObjectId Id, string Layer)>();

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
                    }
                    catch { /* proxy or locked entities */ }
                }

                result.TotalCircles = allCircles.Count;
                result.TotalTexts   = allTexts.Count;

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
                    result.Log += $"Wykryto istniejący styl PH: warstwa={detectedLayer}, styl={detectedStyle}\n";
                }

                // Ensure annotation layers and text style exist
                EnsureLayer(tr, db, detectedLayer, 2);   // yellow
                EnsureLayer(tr, db, LayerPhHatch, 256);  // byLayer
                string styleId = EnsureTextStyle(tr, db, detectedStyle);

                btr.UpgradeOpen();

                // ── 3. Annotate each pile ─────────────────────────────────────────────
                foreach (var pile in piles)
                {
                    // NO ACTION → no annotation
                    if (pile.PhAction == "NO ACTION" ||
                        string.IsNullOrEmpty(pile.PhAction))
                    {
                        result.Skipped.Add(pile.PileId);
                        continue;
                    }

                    // Find text entity matching this pile ID
                    var matchText = FindPileText(pile.PileId, allTexts);
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

                    if (nearCircle.Id == ObjectId.Null)
                    {
                        result.NotFound.Add($"{pile.PileId}(no circle)");
                        continue;
                    }

                    double radius  = nearCircle.Radius;
                    Point3d center = nearCircle.Center;

                    // ── Add SOLID hatch inside circle ─────────────────────────────
                    AddSolidHatch(tr, db, btr, center, radius, pile.PhAction);

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

                tr.Commit();
            }

            result.Log += BuildLog(result);
            return result;
        }

        // ── Private helpers ──────────────────────────────────────────────────────────

        private static void AddSolidHatch(Transaction tr, Database db,
            BlockTableRecord btr, Point3d center, double radius, string phAction)
        {
            // Boundary circle (invisible)
            var boundCircle = new Circle(center, Vector3d.ZAxis, radius);
            boundCircle.SetDatabaseDefaults();
            boundCircle.Layer   = LayerPhHatch;
            boundCircle.Visible = false;
            btr.AppendEntity(boundCircle);
            tr.AddNewlyCreatedDBObject(boundCircle, true);

            // ANSI31 diagonal hatch — Red (ACI 1)
            // Scale proporcjonalny do promienia: ~5-6 linii przez średnicę
            // ANSI31 base spacing = 3.175mm @ scale=1 → scale = radius * 0.12 ≈ spacing co ~38mm przy r=300
            double hatchScale = Math.Max(5.0, radius * 0.12);

            var hatch = new Hatch();
            hatch.SetDatabaseDefaults();
            hatch.HatchObjectType = HatchObjectType.HatchObject;
            hatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI31");
            hatch.PatternScale = hatchScale;
            hatch.ColorIndex   = 1;   // red
            hatch.Layer        = LayerPhHatch;
            btr.AppendEntity(hatch);
            tr.AddNewlyCreatedDBObject(hatch, true);
            hatch.Associative = false;
            hatch.AppendLoop(HatchLoopTypes.Outermost, new ObjectIdCollection { boundCircle.ObjectId });
            hatch.EvaluateHatch(true);
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
            FindPileText(string pileId, List<(string Text, Point3d Pos, ObjectId Id, double Height, string Style, string Layer)> texts)
        {
            string normId = NormalizePileId(pileId);
            foreach (var t in texts)
                if (NormalizePileId(t.Text) == normId)
                    return t;
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
            // Extract number from phAction (PH3, PH3-RE, EXCEED, etc.)
            var m = Regex.Match(phAction, @"\d+(-\w+)?");
            string suffix = m.Success ? m.Value : phAction.Replace("PH", "");

            if (phAction == "EXCEED")
                return @"{\CEXCEED}"; // fallback plain text

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
                case "PH3-RE":                      return 6;   // magenta
                case "EXCEED":                      return 10;  // dark red / maroon
                default:                            return 7;   // white
            }
        }

        private static string BuildLog(AnnotationResult r)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Podpisano: {r.Annotated.Count} | Pominięto NO ACTION: {r.Skipped.Count} | Nie znaleziono: {r.NotFound.Count}");
            if (r.NotFound.Count > 0)
                sb.AppendLine($"Nie znaleziono: {string.Join(", ", r.NotFound)}");
            return sb.ToString();
        }
    }

    public class AnnotationResult
    {
        public int          TotalCircles { get; set; }
        public int          TotalTexts   { get; set; }
        public List<string> Annotated    { get; set; } = new List<string>();
        public List<string> Skipped      { get; set; } = new List<string>();
        public List<string> NotFound     { get; set; } = new List<string>();
        public string       Log          { get; set; } = "";
    }
}
