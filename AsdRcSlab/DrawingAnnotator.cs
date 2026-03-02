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
    /// Annotates pile circles on the active AutoCAD drawing:
    /// - finds CIRCLE entities near TEXT entities matching each pile ID
    /// - adds a TEXT with the PH designation (e.g., PH3)
    /// - fills the circle with a SOLID hatch coloured by PH level
    /// </summary>
    public static class DrawingAnnotator
    {
        // Layer names created by this tool (public so Commands.cs can reference them)
        public const string LayerPhText  = "SD-PH";
        public const string LayerPhHatch = "SD-PH-HATCH";

        public static AnnotationResult Annotate(List<PileData> piles)
        {
            var result = new AnnotationResult();

            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                result.Log = "Brak aktywnego dokumentu AutoCAD.";
                return result;
            }

            var db = doc.Database;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Ensure annotation layers exist
                EnsureLayer(tr, db, LayerPhText,  2);   // yellow
                EnsureLayer(tr, db, LayerPhHatch, 256); // byLayer (colour set per PH)

                var bt  = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // Collect all TEXT / MTEXT / CIRCLE entities from model space
                var allTexts   = new List<(string Text, Point3d Pos, ObjectId Id)>();
                var allCircles = new List<(Point3d Center, double Radius, ObjectId Id)>();

                foreach (ObjectId id in btr)
                {
                    try
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead);
                        if (ent is DBText dbt)
                            allTexts.Add((dbt.TextString?.Trim() ?? "", dbt.Position, id));
                        else if (ent is MText mt)
                        {
                            // Strip MText formatting to get plain text
                            string plain = mt.Contents?.Replace("\\P", " ") ?? "";
                            plain = Regex.Replace(plain, @"\\\w[^;]*;", "");
                            plain = plain.Trim();
                            allTexts.Add((plain, mt.Location, id));
                        }
                        else if (ent is Circle c)
                            allCircles.Add((c.Center, c.Radius, id));
                    }
                    catch { /* ignore locked / proxy entities */ }
                }

                result.TotalCircles = allCircles.Count;
                result.TotalTexts   = allTexts.Count;

                // Switch btr to write for appending new entities
                btr.UpgradeOpen();

                foreach (var pile in piles)
                {
                    // Find text entity that best matches this pile ID
                    var match = FindPileText(pile.PileId, allTexts);
                    if (match == null)
                    {
                        result.NotFound.Add(pile.PileId);
                        continue;
                    }

                    Point3d textPos = match.Value.Pos;

                    // Find nearest circle to the pile text
                    var nearCircle = allCircles
                        .Where(c => c.Center.DistanceTo(textPos) < c.Radius * 10)
                        .OrderBy(c => c.Center.DistanceTo(textPos))
                        .FirstOrDefault();

                    if (nearCircle.Id == ObjectId.Null)
                    {
                        result.NotFound.Add($"{pile.PileId}(no circle)");
                        continue;
                    }

                    double r = nearCircle.Radius;
                    Point3d cen = nearCircle.Center;

                    // Position PH text below the pile ID text (offset = 1.5 × radius)
                    double textHeight = Math.Max(r * 0.5, 50);
                    var phPos = new Point3d(textPos.X, textPos.Y - r * 1.5, 0);

                    // Add PH label TEXT
                    var phText = new DBText();
                    phText.SetDatabaseDefaults();
                    phText.TextString = pile.PhAction;
                    phText.Height     = textHeight;
                    phText.Position   = phPos;
                    phText.Layer      = LayerPhText;
                    phText.Justify    = AttachmentPoint.MiddleCenter;
                    phText.AlignmentPoint = phPos;
                    btr.AppendEntity(phText);
                    tr.AddNewlyCreatedDBObject(phText, true);

                    // Add SOLID HATCH inside the circle
                    var hatch = new Hatch();
                    hatch.SetDatabaseDefaults();
                    hatch.HatchObjectType = HatchObjectType.HatchObject;
                    hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                    hatch.ColorIndex = PhColorIndex(pile.PhAction);
                    hatch.Layer      = LayerPhHatch;
                    btr.AppendEntity(hatch);
                    tr.AddNewlyCreatedDBObject(hatch, true);

                    // Boundary: create a duplicate circle as hatch boundary source
                    var boundaryCopy = new Circle(cen, Vector3d.ZAxis, r);
                    boundaryCopy.SetDatabaseDefaults();
                    boundaryCopy.Layer   = LayerPhHatch;
                    boundaryCopy.Visible = false;
                    btr.AppendEntity(boundaryCopy);
                    tr.AddNewlyCreatedDBObject(boundaryCopy, true);

                    var idCol = new ObjectIdCollection { boundaryCopy.ObjectId };
                    hatch.Associative = false;
                    hatch.AppendLoop(HatchLoopTypes.Outermost, idCol);
                    hatch.EvaluateHatch(true);

                    result.Annotated.Add(pile.PileId);
                }

                tr.Commit();
            }

            result.Log = BuildLog(result);
            return result;
        }

        // ── Helpers ─────────────────────────────────────────────────────────────────

        private static (string Text, Point3d Pos, ObjectId Id)? FindPileText(
            string pileId, List<(string Text, Point3d Pos, ObjectId Id)> allTexts)
        {
            string normId = NormalizePileId(pileId);

            foreach (var t in allTexts)
            {
                if (NormalizePileId(t.Text) == normId)
                    return t;
            }
            return null;
        }

        /// <summary>Normalizuje pile ID: usuwa prefix P/p, leading zeros, trim.</summary>
        private static string NormalizePileId(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim().TrimStart('P', 'p');
            if (int.TryParse(s, out int n)) return n.ToString();
            return s.ToUpperInvariant();
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

        // ACI color index by PH level
        private static short PhColorIndex(string ph)
        {
            switch (ph)
            {
                case "PH1": case "PH2": case "PH3": return 3;  // green
                case "PH4": case "PH5": case "PH6": return 2;  // yellow
                case "PH7": case "PH8": case "PH9": return 1;  // red
                case "PH3-RE":                      return 6;  // magenta
                case "EXCEED":                      return 10; // dark red
                default:                            return 7;  // white
            }
        }

        private static string BuildLog(AnnotationResult r)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Podpisano: {r.Annotated.Count} pali");
            if (r.NotFound.Count > 0)
                sb.AppendLine($"Nie znaleziono: {string.Join(", ", r.NotFound)}");
            return sb.ToString();
        }
    }

    public class AnnotationResult
    {
        public int         TotalCircles { get; set; }
        public int         TotalTexts   { get; set; }
        public List<string> Annotated   { get; set; } = new List<string>();
        public List<string> NotFound    { get; set; } = new List<string>();
        public string      Log          { get; set; } = "";
    }
}
