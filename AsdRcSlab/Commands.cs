using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Microsoft.Win32;

namespace AsdRcSlab
{
    public class Commands
    {
        private const string TitleBlockName    = "A1-BL";
        private const string GaSlabNotesLayer  = "PCN-Text";
        private const string RcSlabNotesLayer  = "SD-Text";
        private const string RcnRefsLayer      = "ASD-RCN-REFS";
        private const string KeySlabArea             = "SLAB_AREA";
        private const string KeySlabPerimeter        = "SLAB_PERIMETER";
        private const string KeySlabThickness        = "SLAB_THICKNESS";
        private const string KeyConcreteVolume       = "CONCRETE_VOLUME";
        private const string KeyConcreteDesignated   = "CONCRETE_DESIGNATED_BLOCK";

        private const string KeySlabAreaRaw       = "area_raw";
        private const string KeySlabPerimeterRaw  = "perimeter_raw";
        private const string KeySlabThicknessRaw  = "thickness_raw";
        private const string KeyConcreteVolumeRaw = "volume_raw";

        private static readonly string[] GaiFieldsToCopy = new[]
        {
            "CLIENT_1", "CLIENT_2", "CLIENT_3",
            "PROJ_1",   "PROJ_2",   "PROJ_3",
            "APPROVED",
        };

        // Slab extract: group 1 = numeric value
        private static readonly Regex SlabAreaExtractRx      = new Regex(@"SLAB\s+AREA\s*=\s*([\d.]+)\s*m",       RegexOptions.IgnoreCase);
        private static readonly Regex SlabPerimeterExtractRx = new Regex(@"SLAB\s+PERIMETER\s*=\s*([\d.]+)\s*m",  RegexOptions.IgnoreCase);
        private static readonly Regex SlabThicknessExtractRx = new Regex(@"SLAB\s+THICKNESS\s*=\s*([\d.]+)\s*mm", RegexOptions.IgnoreCase);

        // CONCRETE VOLUME extract — g1=prefix, g2=number only (tail ignorowany)
        private static readonly Regex ConcreteVolumeExtractRx = new Regex(
            @"(CONCRETE\s+VOLUME\s*=\s*)([\d.]+)",
            RegexOptions.IgnoreCase);

        // Raw fallback extract: capture all after "=" to first MText control code or end
        private static readonly Regex SlabAreaRawExtractRx = new Regex(
            @"SLAB\s+AREA\s*=\s*([^\\]+?)(?=\\|$)",
            RegexOptions.IgnoreCase);
        private static readonly Regex SlabPerimeterRawExtractRx = new Regex(
            @"SLAB\s+PERIMETER\s*=\s*([^\\]+?)(?=\\|$)",
            RegexOptions.IgnoreCase);
        private static readonly Regex SlabThicknessRawExtractRx = new Regex(
            @"SLAB\s+THICKNESS\s*=\s*([^\\]+?)(?=\\|$)",
            RegexOptions.IgnoreCase);
        private static readonly Regex ConcreteVolumeRawExtractRx = new Regex(
            @"CONCRETE\s+VOLUME\s*=\s*([^\\]+?)(?=\\|$)",
            RegexOptions.IgnoreCase);

        // CONCRETE VOLUME replace w RC: G1=prefix, G2=old value (numeric or MText-formatted), G3=m³ codes z RC (zachowane); stary tail wycinany
        private static readonly Regex ConcreteVolumeReplaceRx = new Regex(
            @"(CONCRETE\s+VOLUME\s*=\s*)((?:[^\\]|\\[A-Za-z][^;\\]*;)*?)(\s*m(?:\\[A-Za-z][^;]*;|\s)*[³3]?(?:\\[A-Za-z][^;]*;|\s)*)[^\\]*?(?=\\P|\d+\.\s|\z)",
            RegexOptions.IgnoreCase);

        // Podmienia HYSTOOLS DK90/DK165 (w tym samym MText co SLAB NOTES)
        private static readonly Regex HystoolsRx = new Regex(
            @"HYSTOOLS\s+DK(?:90|165)",
            RegexOptions.IgnoreCase);

        // Slab replace: group 1 = "X = ", group 2 = old value (numeric or MText-formatted), group 3 = " m"/" mm"
        private static readonly Regex SlabAreaReplaceRx = new Regex(
            @"(SLAB\s+AREA\s*=\s*)((?:[^\\]|\\[A-Za-z][^;\\]*;)*?)(\s*m)(?=\\|$)",
            RegexOptions.IgnoreCase);
        private static readonly Regex SlabPerimeterReplaceRx = new Regex(
            @"(SLAB\s+PERIMETER\s*=\s*)((?:[^\\]|\\[A-Za-z][^;\\]*;)*?)(\s*m)(?=\\|$)",
            RegexOptions.IgnoreCase);
        private static readonly Regex SlabThicknessReplaceRx = new Regex(
            @"(SLAB\s+THICKNESS\s*=\s*)((?:[^\\]|\\[A-Za-z][^;\\]*;)*?)(\s*mm)(?=\\|$)",
            RegexOptions.IgnoreCase);

        // Raw replace: replaces all content after "=" (to first backslash or end) with raw value
        private static readonly Regex SlabAreaRawReplaceRx = new Regex(
            @"(SLAB\s+AREA\s*=\s*)[^\\]*",
            RegexOptions.IgnoreCase);
        private static readonly Regex SlabPerimeterRawReplaceRx = new Regex(
            @"(SLAB\s+PERIMETER\s*=\s*)[^\\]*",
            RegexOptions.IgnoreCase);
        private static readonly Regex SlabThicknessRawReplaceRx = new Regex(
            @"(SLAB\s+THICKNESS\s*=\s*)[^\\]*",
            RegexOptions.IgnoreCase);
        private static readonly Regex ConcreteVolumeRawReplaceRx = new Regex(
            @"(CONCRETE\s+VOLUME\s*=\s*)[^\\]*",
            RegexOptions.IgnoreCase);

        // Universal fallback — brak wymagania unit suffix; lookahead na \P, text suffix, lub koniec
        private static readonly Regex SlabAreaUniversalRx = new Regex(
            @"(SLAB\s+AREA\s*=\s*)(?:[^\\]|\\[A-Za-z][^;\\]*;)*?(?=\\P|$)",
            RegexOptions.IgnoreCase);
        private static readonly Regex SlabPerimeterUniversalRx = new Regex(
            @"(SLAB\s+PERIMETER\s*=\s*)(?:[^\\]|\\[A-Za-z][^;\\]*;)*?(?=\\P|$)",
            RegexOptions.IgnoreCase);
        private static readonly Regex SlabThicknessUniversalRx = new Regex(
            @"(SLAB\s+THICKNESS\s*=\s*)(?:[^\\]|\\[A-Za-z][^;\\]*;)*?(?=\\P|\s+U\.N\.O\.|$)",
            RegexOptions.IgnoreCase);
        private static readonly Regex ConcreteVolumeUniversalRx = new Regex(
            @"(CONCRETE\s+VOLUME\s*=\s*)(?:[^\\]|\\[A-Za-z][^;\\]*;)*?(?=\\P|\s+PILE\s+CAP\s+INC\.|$)",
            RegexOptions.IgnoreCase);

        // CONCRETE TO BE DESIGNATED block — od frazy do "CERTIFICATE." (włącznie z kropką)
        private static readonly Regex ConcreteDesignatedBlockRx = new Regex(
            @"CONCRETE\s+TO\s+BE\s+DESIGNATED.*?CERTIFICATE\s*\.",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        // Wyciąga numer PLOT z TITLE_1 (np. "PLOT 3" → 3)
        private static readonly Regex TitlePlotNumberRx =
            new Regex(@"\bPLOT\s+(\d+)", RegexOptions.IgnoreCase);

        // Wyciąga numer z suffix GA (np. "GA100" → 100)
        private static readonly Regex GaSuffixNumberRx =
            new Regex(@"^GA(\d+)$", RegexOptions.IgnoreCase);

        // Auto-detekcja TITLE_3
        private static readonly Regex MainLayerRx = new Regex(
            @"REINFORCEMENT\s+DETAILS\s+.*?\b(BOTTOM|TOP)\s+LAYER",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        private static readonly Regex SectionRx = new Regex(
            @"\bSECTION\s+[A-Z]\s*-\s*[A-Z]\b",
            RegexOptions.IgnoreCase);

        private static readonly Regex PhRx = new Regex(
            @"\bPH[1-9](-RE)?\b",
            RegexOptions.IgnoreCase);

        private static readonly Regex DetailRx = new Regex(
            @"\bDETAIL\s+['""]?\d+['""]?",
            RegexOptions.IgnoreCase);

        // ── PANEL 1: PROJEKT ──────────────────────────────────────────────────

        [CommandMethod("ASD-PROJ")]
        public void CmdNoweProjekt()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var dlg = new NewProjectDialog(SessionData.CurrentProject);
            if (AcApp.ShowModalWindow(AcApp.MainWindow.Handle, dlg, false) == true)
            {
                SessionData.CurrentProject = dlg.Result;

                // Zapisz project.json obok aktywnego DWG
                string dwgPath = doc.Database.Filename;
                string folder  = string.IsNullOrEmpty(dwgPath)
                    ? System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
                    : Path.GetDirectoryName(dwgPath);

                string jsonPath = Path.Combine(folder, "project.json");
                File.WriteAllText(jsonPath, JsonConvert.SerializeObject(dlg.Result, Formatting.Indented));

                doc.Editor.WriteMessage($"\nProjekt zapisany: {dlg.Result.DRWNumber} → {jsonPath}\n");
            }
        }

        [CommandMethod("ASD-OPEN")]
        public void CmdOtworzProjekt()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;

            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title  = "Open Project",
                Filter = "Project JSON (*.json)|*.json",
                DefaultExt = ".json"
            };

            if (dlg.ShowDialog() != true) return;

            try
            {
                string json = File.ReadAllText(dlg.FileName);
                SessionData.CurrentProject = JsonConvert.DeserializeObject<ProjectData>(json);
                doc.Editor.WriteMessage(
                    $"\nLoaded: {SessionData.CurrentProject.ProjectName}" +
                    $", DRW: {SessionData.CurrentProject.DRWNumber}" +
                    $", h={SessionData.CurrentProject.SlabThickness}mm\n");
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\nError loading project: {ex.Message}\n");
            }
        }

        [CommandMethod("ASD-GAI")]
        public void GaCopyAttributes()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            // 1. File picker
            var dlg = new OpenFileDialog
            {
                Title  = "Select GA file (title block source)",
                Filter = "AutoCAD drawings (*.dwg;*.dxf)|*.dwg;*.dxf"
            };
            if (dlg.ShowDialog() != true) return;
            string path = dlg.FileName;
            ed.WriteMessage($"\nGAI: reading GA from '{Path.GetFileName(path)}'...");

            // 2. Czytaj atrybuty A1-BL (zawiera też TITLE_1)
            Dictionary<string, string> srcAttrs;
            try
            {
                srcAttrs = ReadA1BLAttributesFromFile(path);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Cannot open GA file:\n{ex.Message}",
                                "GAI", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (srcAttrs == null || srcAttrs.Count == 0)
            {
                MessageBox.Show($"Title block not found in GA file: '{TitleBlockName}'.",
                                "GAI", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 3. Wyciągnij TITLE prefix i DRAWING_NUMBER prefix z już wczytanego A1-BL
            string gaTitlePrefix   = null;
            string gaDrawingPrefix = null;

            if (srcAttrs.TryGetValue("TITLE_1", out var gaTitle1Raw))
                gaTitlePrefix = ExtractGaTitlePrefix(gaTitle1Raw);

            if (srcAttrs.TryGetValue("DRAWING_NUMBER", out var gaDrawNo))
            {
                int dashIdx = gaDrawNo.IndexOf('-');
                if (dashIdx > 0) gaDrawingPrefix = gaDrawNo.Substring(0, dashIdx);
            }

            // 4. Otwórz GA drugi raz dla SLAB values (MText na PCN-Text) + first GA number
            var gaSlabValues = new Dictionary<string, string>();
            int? firstGaNumber = null;
            try
            {
                using (var sideDb = new Database(false, true))
                {
                    bool isDxf = path.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase);
                    if (isDxf) sideDb.DxfIn(path, null);
                    else       sideDb.ReadDwgFile(path, System.IO.FileShare.Read, true, "");

                    gaSlabValues = ExtractSlabValuesFromDb(sideDb, GaSlabNotesLayer);

                    try
                    {
                        firstGaNumber = ExtractFirstGaDrawingNumber(sideDb);
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nGAI: failed to read first GA number: {ex.Message}");
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error reading GA (SLAB notes):\n{ex.Message}",
                                "GAI", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 5. Wyciągnij numer plotu z GA TITLE_1 prefix (np. "PLOT 58." → 58)
            int? plotNumberFromGa = null;
            if (!string.IsNullOrEmpty(gaTitlePrefix))
            {
                var pm = TitlePlotNumberRx.Match(gaTitlePrefix);
                if (pm.Success && int.TryParse(pm.Groups[1].Value, out int p))
                    plotNumberFromGa = p;
            }

            // 6. Preview + confirm
            var sb = new StringBuilder();
            sb.AppendLine("Copy the following values to RC?");
            sb.AppendLine();
            sb.AppendLine("PROJECT DATA:");
            foreach (var tag in GaiFieldsToCopy)
            {
                string val = srcAttrs.TryGetValue(tag, out var v) && !string.IsNullOrEmpty(v) ? v : "(none)";
                sb.AppendLine($"  {tag,-10}: {val}");
            }
            sb.AppendLine();
            sb.AppendLine("TITLE_1:");
            if (string.IsNullOrEmpty(gaTitlePrefix))
            {
                sb.AppendLine("  GA prefix: (none)");
                sb.AppendLine("  → RC TITLE_1: prefix removed");
            }
            else
            {
                sb.AppendLine($"  GA prefix: \"{gaTitlePrefix}\"");
                sb.AppendLine($"  → RC TITLE_1: \"{gaTitlePrefix} <rest after prefix>\"");
            }
            sb.AppendLine();
            sb.AppendLine("DRAWING_NUMBER:");
            if (string.IsNullOrEmpty(gaDrawingPrefix))
            {
                sb.AppendLine("  (no GA prefix — won't overwrite)");
            }
            else
            {
                string startInfo;
                if (firstGaNumber.HasValue)
                    startInfo = $"Starting from GA: GA{firstGaNumber.Value} → RC{firstGaNumber.Value}, RC{firstGaNumber.Value + 1}, ...";
                else if (plotNumberFromGa.HasValue)
                    startInfo = $"Plot {plotNumberFromGa.Value} (from TITLE_1 prefix, fallback)";
                else
                    startInfo = "No GA number or plot, fallback RC001+";
                sb.AppendLine($"  {startInfo}");

                var rcLayoutNames = new List<string>();
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var ld = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                    foreach (DBDictionaryEntry ent in ld)
                    {
                        var lay = tr.GetObject(ent.Value, OpenMode.ForRead) as Layout;
                        if (lay == null) continue;
                        if (string.Equals(lay.LayoutName, "Model", StringComparison.OrdinalIgnoreCase)) continue;
                        rcLayoutNames.Add(lay.LayoutName);
                    }
                    tr.Commit();
                }
                rcLayoutNames.Sort(StringComparer.OrdinalIgnoreCase);

                for (int i = 0; i < rcLayoutNames.Count; i++)
                {
                    string suffix = BuildDrawingSuffix(plotNumberFromGa, i, firstGaNumber);
                    sb.AppendLine($"  {rcLayoutNames[i]} → {gaDrawingPrefix}-{suffix}");
                }
            }
            sb.AppendLine();
            sb.AppendLine("SLAB NOTES (numeric values):");
            sb.AppendLine($"  SLAB AREA       : {(gaSlabValues.TryGetValue(KeySlabArea,       out var sa)  && !string.IsNullOrEmpty(sa)  ? sa  + " m²" : "(none)")}");
            sb.AppendLine($"  SLAB PERIMETER  : {(gaSlabValues.TryGetValue(KeySlabPerimeter,  out var spe) && !string.IsNullOrEmpty(spe) ? spe + " m"   : "(none)")}");
            sb.AppendLine($"  SLAB THICKNESS  : {(gaSlabValues.TryGetValue(KeySlabThickness,  out var sth) && !string.IsNullOrEmpty(sth) ? sth + " mm"  : "(none)")}");
            sb.AppendLine($"  CONCRETE VOLUME = {(gaSlabValues.TryGetValue(KeyConcreteVolume, out var svo) && !string.IsNullOrEmpty(svo) ? svo + "m³"   : "(none)")}");
            sb.AppendLine();
            if (gaSlabValues.TryGetValue(KeySlabThickness, out var sthHys) &&
                int.TryParse(sthHys, out int thHys))
            {
                string targetDk = thHys == 225 ? "DK90"
                               : thHys == 300 ? "DK165"
                               : "—";
                sb.AppendLine($"  HYSTOOLS (from thickness {thHys}mm): → {targetDk}");
            }
            else
            {
                sb.AppendLine("  HYSTOOLS: thickness not detected — no substitution");
            }
            sb.AppendLine();
            sb.AppendLine("CONCRETE TO BE DESIGNATED block:");
            if (gaSlabValues.TryGetValue(KeyConcreteDesignated, out var dBlock) && !string.IsNullOrEmpty(dBlock))
            {
                string clean = Regex.Replace(dBlock, @"\\[A-Za-z][^;\s]*;?", " ").Replace("\\P", " ");
                clean = Regex.Replace(clean, @"\s+", " ").Trim();
                if (clean.Length > 200) clean = clean.Substring(0, 197) + "...";
                sb.AppendLine($"  {clean}");
            }
            else
            {
                sb.AppendLine("  (none — not found in GA)");
            }

            var confirmResult = MessageBox.Show(sb.ToString(), "GAI — confirm",
                                                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirmResult != MessageBoxResult.Yes)
            {
                ed.WriteMessage("\nGAI: cancelled by user.");
                return;
            }

            // 6. Apply: atrybuty A1-BL + TITLE_1 prefix
            int updatedAttrLayouts;
            try
            {
                updatedAttrLayouts = ApplyA1BLAttributesToActiveDb(
                    db, srcAttrs, GaiFieldsToCopy,
                    gaTitlePrefix, gaDrawingPrefix, plotNumberFromGa, firstGaNumber);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error overwriting attributes:\n{ex.Message}",
                                "GAI", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 7. Apply: SLAB NOTES wartości
            int updatedSlabLayouts = 0;
            if (gaSlabValues.Count > 0)
            {
                try
                {
                    updatedSlabLayouts = ApplySlabValuesToActiveDb(db, gaSlabValues, RcSlabNotesLayer);
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nGAI: SLAB warning: {ex.Message}");
                }
            }

            // 8. Log + komunikat
            ed.WriteMessage($"\nGAI: updated {updatedAttrLayouts} layout(s) with title block, " +
                            $"{updatedSlabLayouts} layout(s) with SLAB NOTES.");
            MessageBox.Show($"Updated:\n• title block: {updatedAttrLayouts} layout(s)\n" +
                            $"• SLAB NOTES: {updatedSlabLayouts} layout(s)",
                            "GAI", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        [CommandMethod("ASD-RCN")]
        public void RcAutoNaming()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            // 1. Wyciąganie auto-wartości
            Dictionary<string, string> autoTitle3Map;
            Dictionary<string, string> autoScalesMap;

            try
            {
                autoTitle3Map = ExtractAutoTitle3(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nRCN: TITLE_3 auto-detection error: {ex.Message}");
                autoTitle3Map = new Dictionary<string, string>();
            }

            try
            {
                autoScalesMap = ExtractLayoutScales(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nRCN: SCALE auto-detection error: {ex.Message}");
                autoScalesMap = new Dictionary<string, string>();
            }

            string nowDate = DateTime.Now.ToString(
                "MMM yyyy",
                System.Globalization.CultureInfo.InvariantCulture).ToUpper();

            // 2. Preview
            var sb = new StringBuilder();
            sb.AppendLine("ASD-RCN — auto-fill RC title blocks.");
            sb.AppendLine();

            sb.AppendLine("TITLE_3 (auto from Model space + viewports):");
            if (autoTitle3Map.Count == 0)
            {
                sb.AppendLine("  (none — nothing detected)");
            }
            else
            {
                foreach (var kv in autoTitle3Map.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
                {
                    string val = string.IsNullOrEmpty(kv.Value)
                        ? "(not detected — keeping existing)"
                        : "\"" + kv.Value + "\"";
                    sb.AppendLine($"  {kv.Key} → {val}");
                }
            }

            sb.AppendLine();
            sb.AppendLine("SCALE (auto from viewports):");
            if (autoScalesMap.Count == 0)
            {
                sb.AppendLine("  (none)");
            }
            else
            {
                foreach (var kv in autoScalesMap.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
                {
                    string val = string.IsNullOrEmpty(kv.Value)
                        ? "(no viewports)"
                        : "\"" + kv.Value + "\"";
                    sb.AppendLine($"  {kv.Key} → {val}");
                }
            }

            sb.AppendLine();
            sb.AppendLine($"DATE: {nowDate}");

            sb.AppendLine();
            sb.AppendLine("LAYOUT RENAME (after update — by DRAWING_NUMBER suffix + 'C1'):");
            sb.AppendLine("  (e.g. SL44QR001-RC030 → RC030C1)");

            sb.AppendLine();
            sb.AppendLine("REFERENCES \"SEE DRG\":");
            sb.AppendLine($"  Each layout gets frames to all others.");
            sb.AppendLine($"  Layer: {RcnRefsLayer} (created if not exists).");
            sb.AppendLine($"  Position: x=900, y from 580 down (outside A1 sheet).");
            sb.AppendLine($"  Existing frames on this layer will be replaced.");

            sb.AppendLine();
            sb.AppendLine("Apply?");

            var result = MessageBox.Show(
                sb.ToString(), "ASD-RCN",
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
            {
                ed.WriteMessage("\nRCN: cancelled.");
                return;
            }

            // 3. Apply
            int updatedLayouts = 0;
            try
            {
                updatedLayouts = ApplyAutoFieldsToActiveDb(db, autoTitle3Map, autoScalesMap, nowDate);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error saving attributes:\n{ex.Message}",
                                "RCN", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 4. Rename layoutów
            int renamedLayouts = 0;
            try
            {
                renamedLayouts = RenameLayoutsFromDrawingNumber(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nRCN: layout rename failed: {ex.Message}");
            }

            // 5. Stwórz ramki z odnośnikami "SEE DRG ..."
            int refFrames = 0;
            try
            {
                refFrames = CreateReferenceFramesOnAllLayouts(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nRCN: error creating reference frames: {ex.Message}");
                ed.WriteMessage($"\nRCN REFS EXCEPTION: {ex}");
            }

            // 6. Podsumowanie
            string summary = $"ASD-RCN — done.\n\n" +
                             $"Updated attributes in {updatedLayouts} layout(s).\n" +
                             $"Renamed {renamedLayouts} layout(s).\n" +
                             $"Inserted {refFrames} reference frames (layer {RcnRefsLayer}).";
            MessageBox.Show(summary, "ASD-RCN", MessageBoxButton.OK, MessageBoxImage.Information);
            ed.WriteMessage($"\nRCN: updated {updatedLayouts}, renamed {renamedLayouts}, refs {refFrames}.");
        }

        [CommandMethod("ASD-SET")]
        public void CmdUstawienia()
        {
            var dlg = new SettingsDialog();
            AcApp.ShowModalWindow(AcApp.MainWindow.Handle, dlg, false);
        }

        // ── PANEL 2: ZBROJENIE ────────────────────────────────────────────────

        [CommandMethod("ASD-GBOT")]
        public void CmdGenerujB1B2()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed  = doc.Editor;

            var peo = new PromptEntityOptions("\nSelect slab outline (layer SD-PILED-RAFT): ");
            peo.SetRejectMessage("\nSelect entity.");
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            string validationError = null;
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var ent = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                if (ent == null)
                    validationError = "Cannot open entity.";
                else if (ent.Layer != "SD-PILED-RAFT")
                    validationError = "Select polyline on layer SD-PILED-RAFT.";
                else if (!(ent is Polyline))
                    validationError = "Selected entity is not a polyline (LWPolyline).";
                tr.Commit();
            }

            if (validationError != null)
            {
                System.Windows.MessageBox.Show(validationError, "ASD-GBOT",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            if (SessionData.TemplateBarsB.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    "First create H10 bars (rbcr_def_bar_bv) and register them with ASD-GSETUP.",
                    "ASD-GBOT", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            var result = ReinforcementGenerator.GenerateBottomAsd(doc, per.ObjectId,
                SessionData.TemplateBarsB);
            if (!string.IsNullOrEmpty(result.Error))
            {
                ed.WriteMessage($"\nGBOT error: {result.Error}\n");
                System.Windows.MessageBox.Show(result.Error, "ASD-GBOT — error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            SessionData.LapPositionsB1 = result.LapPositionsX;
            SessionData.LapPositionsB2 = result.LapPositionsY;
            ed.WriteMessage($"\nB1/B2: sending {result.BarsDrawn} bars to ASD...\n");
        }

        [CommandMethod("ASD-GTOP")]
        public void CmdGenerujT1T2()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed  = doc.Editor;

            var peo = new PromptEntityOptions("\nSelect slab outline (layer SD-PILED-RAFT): ");
            peo.SetRejectMessage("\nSelect entity.");
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            string validationError = null;
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var ent = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                if (ent == null)
                    validationError = "Cannot open entity.";
                else if (ent.Layer != "SD-PILED-RAFT")
                    validationError = "Select polyline on layer SD-PILED-RAFT.";
                else if (!(ent is Polyline))
                    validationError = "Selected entity is not a polyline (LWPolyline).";
                tr.Commit();
            }

            if (validationError != null)
            {
                System.Windows.MessageBox.Show(validationError, "ASD-GTOP",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            if (SessionData.TemplateBarsT.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    "First create H12 bars (rbcr_def_bar_bv) and register them with ASD-GSETUP.",
                    "ASD-GTOP", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            var result = ReinforcementGenerator.GenerateTopAsd(doc, per.ObjectId,
                SessionData.TemplateBarsT,
                SessionData.LapPositionsB1, SessionData.LapPositionsB2);
            if (!string.IsNullOrEmpty(result.Error))
            {
                ed.WriteMessage($"\nGTOP error: {result.Error}\n");
                System.Windows.MessageBox.Show(result.Error, "ASD-GTOP — error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            ed.WriteMessage($"\nT1/T2: sending {result.BarsDrawn} bars to ASD...\n");
        }

        [CommandMethod("ASD-BMM")]
        public void CmdOznaczPrety()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;

            var fileDlg = new Microsoft.Win32.OpenFileDialog
            {
                Title      = "Select BBS file",
                Filter     = "Excel BBS (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx"
            };
            if (fileDlg.ShowDialog() != true) return;

            try
            {
                var result   = BmmChecker.CheckAll(fileDlg.FileName);
                int failCount = new[] { result.R87, result.R95, result.R81, result.R83, result.R92 }
                    .Count(r => r.Status == "FAIL");

                doc.Editor.WriteMessage($"\nBMM: {failCount} error(s) found — check results window.\n");

                var resultDlg = new BmmResultsDialog(result, System.IO.Path.GetFileName(fileDlg.FileName));
                AcApp.ShowModalWindow(AcApp.MainWindow.Handle, resultDlg, false);
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\nBMM error: {ex.Message}\n");
            }
        }

        [CommandMethod("ASD-LAP")]
        public void CmdZakladyAuto()
        {
            var dlg = new LapCalculatorDialog();
            AcApp.ShowModalWindow(AcApp.MainWindow.Handle, dlg, false);
        }

        /// <summary>
        /// One-time setup: user selects all H10 template bars with a window, then all H12.
        /// Bar lengths (1250–6000 mm, step 250) are detected automatically from geometry.
        /// Must be run before ASD-GBOT / ASD-GTOP.
        /// </summary>
        [CommandMethod("ASD-GSETUP")]
        public void CmdGSetup()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed  = doc.Editor;

            // ── H10 bars (B1/B2) ─────────────────────────────────────────────────
            var pso1 = new PromptSelectionOptions();
            pso1.MessageForAdding = "\nSelect all H10 bars in window (B1/B2 layout): ";
            var sel1 = ed.GetSelection(pso1);
            if (sel1.Status != PromptStatus.OK) return;

            var barsB = ReadTemplateBarPositions(doc.Database, sel1.Value.GetObjectIds());
            if (barsB.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    "No H10 bars detected. Make sure ASD bars (rbcr_def_bar_bv) are selected.",
                    "ASD-GSETUP", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }
            SessionData.TemplateBarsB = barsB;

            // ── H12 bars (T1/T2) ─────────────────────────────────────────────────
            var pso2 = new PromptSelectionOptions();
            pso2.MessageForAdding = "\nSelect all H12 bars in window (T1/T2 layout): ";
            var sel2 = ed.GetSelection(pso2);
            if (sel2.Status != PromptStatus.OK) return;

            var barsT = ReadTemplateBarPositions(doc.Database, sel2.Value.GetObjectIds());
            if (barsT.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    "No H12 bars detected.",
                    "ASD-GSETUP", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }
            SessionData.TemplateBarsT = barsT;

            ed.WriteMessage($"\nGSETUP: H10={barsB.Count} bars, H12={barsT.Count} bars. Ready for ASD-GBOT/ASD-GTOP.\n");
            System.Windows.MessageBox.Show(
                $"Templates registered:\n  H10 (B1/B2): {barsB.Count} bars [{string.Join(", ", System.Linq.Enumerable.Select(barsB.Keys, k => k + "mm"))}]\n  H12 (T1/T2): {barsT.Count} bars",
                "ASD-GSETUP", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
        }

        /// <summary>
        /// Reads bar left-endpoints and lengths from a selection of ASD bar entities.
        /// Works for LINE entities (exact) and ASD custom entities (bounding-box approximation).
        /// </summary>
        private static System.Collections.Generic.Dictionary<int, Autodesk.AutoCAD.Geometry.Point3d>
            ReadTemplateBarPositions(Database db, ObjectId[] ids)
        {
            var dict = new System.Collections.Generic.Dictionary<int, Autodesk.AutoCAD.Geometry.Point3d>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var id in ids)
                {
                    try
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;

                        Autodesk.AutoCAD.Geometry.Point3d leftPt;
                        double barLength;

                        if (ent is Line ln)
                        {
                            // Exact endpoints for LINE entities
                            bool startIsLeft = ln.StartPoint.X <= ln.EndPoint.X;
                            leftPt    = startIsLeft ? ln.StartPoint : ln.EndPoint;
                            barLength = ln.Length;
                        }
                        else
                        {
                            // ASD custom entity: use bounding box
                            Extents3d ext;
                            try   { ext = ent.GeometricExtents; }
                            catch { continue; }
                            double w = ext.MaxPoint.X - ext.MinPoint.X;
                            double h = ext.MaxPoint.Y - ext.MinPoint.Y;
                            if (w < h || w < 500) continue; // skip non-horizontal / too short
                            barLength = w;
                            leftPt = new Autodesk.AutoCAD.Geometry.Point3d(
                                ext.MinPoint.X,
                                (ext.MinPoint.Y + ext.MaxPoint.Y) / 2.0, 0);
                        }

                        int lenMm = (int)(System.Math.Round(barLength / 250.0) * 250);
                        if (lenMm < 1000 || lenMm > 7000) continue;

                        if (!dict.ContainsKey(lenMm))
                            dict[lenMm] = leftPt;
                    }
                    catch { }
                }
                tr.Commit();
            }
            return dict;
        }

        // ── PANEL 3: PH CONDITIONS ────────────────────────────────────────────

        [CommandMethod("ASD-PXIE")]
        public void CmdWczytajPunching()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed  = doc.Editor;

            var fileDlg = new Microsoft.Win32.OpenFileDialog
            {
                Title  = "Select Punching file",
                Filter = "Excel (*.xlsx)|*.xlsx"
            };
            if (fileDlg.ShowDialog() != true) return;

            try
            {
                // Krok 1: skanuj ploty z arkusza "Punching Report to Calcs"
                string scanLog;
                var plots = PunchingParser.ScanPlots(fileDlg.FileName, out scanLog);
                ed.WriteMessage($"\n{scanLog}");

                if (plots.Count == 0)
                {
                    System.Windows.MessageBox.Show(
                        "No 'PLOT N' sections found in sheet 'Punching Report to Calcs'.\n\n" +
                        "Check if the file is a punching report in the new format.",
                        "PXIE", System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                    return;
                }

                // Krok 2: wybór plotu (auto jeśli jeden, dialog jeśli wiele)
                PlotInfo selectedPlot;
                if (plots.Count == 1)
                {
                    selectedPlot = plots[0];
                    ed.WriteMessage($"\nPXIE: Auto-selected {selectedPlot}.\n");
                }
                else
                {
                    var dlg = new PlotPickerDialog(plots);
                    if (AcApp.ShowModalWindow(AcApp.MainWindow.Handle, dlg, false) != true)
                        return; // Cancel — SessionData bez zmian
                    selectedPlot = dlg.SelectedPlot;
                }

                // Krok 3: parsuj wybrany plot
                string parseLog;
                var piles = PunchingParser.ParsePlot(fileDlg.FileName, selectedPlot.Number, out parseLog);
                ed.WriteMessage($"\n{parseLog}");

                if (piles.Count == 0)
                {
                    System.Windows.MessageBox.Show(
                        "No piles loaded. Possible reasons:\n" +
                        "• file not recalculated in Excel (open and save Ctrl+S)\n" +
                        "• file sections are empty\n\n" +
                        "Check command line log for details.",
                        "PXIE", System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                    return;
                }

                // Krok 4: zapisz do sesji
                SessionData.Piles       = piles;
                SessionData.CurrentPlot = selectedPlot;
                SessionData.PhAssigned  = false;

                string reentrantPart = selectedPlot.ReentrantCount > 0
                    ? $" REENTRANT:{selectedPlot.ReentrantCount}" : "";
                ed.WriteMessage(
                    $"\nPXIE: Loaded {piles.Count} piles from {selectedPlot} " +
                    $"(INT:{selectedPlot.InternalCount} EDGE:{selectedPlot.EdgeCount} " +
                    $"CORNER:{selectedPlot.CornerCount}{reentrantPart}). Ready for Assign PH.\n");
                System.Windows.MessageBox.Show(
                    $"Loaded {piles.Count} piles from {selectedPlot}.\nReady for Assign PH.",
                    "Load Punching", System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nPXIE error: {ex.Message}\n");
            }
        }

        [CommandMethod("ASD-PAA")]
        public void CmdAssignPH()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;

            if (SessionData.Piles == null || SessionData.Piles.Count == 0)
            {
                doc.Editor.WriteMessage("\nPAA: Run 'Load Punching' (ASD-PXIE) first.\n");
                System.Windows.MessageBox.Show("Load pile data first using 'Load Punching'.",
                    "Assign PH", System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            try
            {
                PhAssigner.AssignAll(SessionData.Piles);
                SessionData.PhAssigned = true;

                doc.Editor.WriteMessage($"\nPAA: PH assigned for {SessionData.Piles.Count} piles.\n");

                // Otwórz dialog z buttonem "Zaktualizuj rysunek" (user decyduje kiedy anotować)
                var dlg = new PhAssignResultsDialog(SessionData.Piles, showUpdateButton: true);
                AcApp.ShowModalWindow(AcApp.MainWindow.Handle, dlg, false);

            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\nPAA error: {ex.Message}\n");
            }
        }

        [CommandMethod("ASD-PHR")]
        public void CmdPHReport()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;

            if (!SessionData.PhAssigned || SessionData.Piles == null)
            {
                doc.Editor.WriteMessage("\nPHR: Run Assign PH first.\n");
                System.Windows.MessageBox.Show("Run Assign PH first.",
                    "PH Report", System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            var dlg = new PhAssignResultsDialog(SessionData.Piles, showUpdateButton: false);
            AcApp.ShowModalWindow(AcApp.MainWindow.Handle, dlg, false);
        }

        [CommandMethod("ASD-PHV")]
        public void CmdWalidujPH()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;

            if (!SessionData.PhAssigned || SessionData.Piles == null)
            {
                doc.Editor.WriteMessage("\nPHV: Run Assign PH first.\n");
                return;
            }

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("PHV — PH VALIDATION:");
            sb.AppendLine(new string('-', 40));

            // R77: brak EXCEED
            var exceed = SessionData.Piles.Where(p => p.PhAction == "EXCEED").ToList();
            if (exceed.Any())
                sb.AppendLine($"R77: FAIL — Util > 100%: {string.Join(", ", exceed.Select(p => p.PileId))}");
            else
                sb.AppendLine("R77: OK — No piles with Util > 100%");

            // R79: brak orphan (puste ApplicablePileIds)
            var orphan = SessionData.Piles.Where(p => p.ApplicablePileIds == null || p.ApplicablePileIds.Count == 0).ToList();
            if (orphan.Any())
                sb.AppendLine($"R79: FAIL — Orphan PH: {string.Join(", ", orphan.Select(p => p.PileId))}");
            else
                sb.AppendLine("R79: OK — All piles have ApplicablePileIds");

            // R27: duplikaty PileId
            var dupes = SessionData.Piles.GroupBy(p => p.PileId).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
            if (dupes.Any())
                sb.AppendLine($"R27: FAIL — Duplikaty: {string.Join(", ", dupes)}");
            else
                sb.AppendLine("R27: OK — No duplicate Pile IDs");

            doc.Editor.WriteMessage($"\n{sb}\n");
            System.Windows.MessageBox.Show(sb.ToString(), "Waliduj PH",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
        }

        // ── GAI helpers ───────────────────────────────────────────────────────

        private static string ExtractGaTitlePrefix(string gaTitle1)
        {
            if (string.IsNullOrWhiteSpace(gaTitle1)) return null;
            var m = Regex.Match(gaTitle1, @"^(.+?)\s+GENERAL\s+ARRANGEMENT",
                                RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value.Trim() : null;
        }

        private static string ReplaceRcTitlePrefix(string rcTitle1, string newPrefix)
        {
            if (string.IsNullOrWhiteSpace(rcTitle1)) return rcTitle1;
            var m = Regex.Match(rcTitle1, @"REINFORCEMENT\s+DETAILS", RegexOptions.IgnoreCase);
            if (!m.Success) return rcTitle1;
            string suffix = rcTitle1.Substring(m.Index);
            if (string.IsNullOrEmpty(newPrefix))
                return suffix;
            else
                return newPrefix + " " + suffix;
        }

        private enum MsTextCategory { MainBottom, MainTop, Section, Ph, Detail }

        private static List<(MsTextCategory cat, double x, double y)> ScanModelSpaceTexts(Database db)
        {
            var result = new List<(MsTextCategory, double, double)>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    string content = null;
                    double x = 0, y = 0;

                    var ent = tr.GetObject(id, OpenMode.ForRead);
                    if (ent is MText mt)
                    {
                        content = mt.Contents;
                        x = mt.Location.X;
                        y = mt.Location.Y;
                    }
                    else if (ent is DBText t)
                    {
                        content = t.TextString;
                        x = t.Position.X;
                        y = t.Position.Y;
                    }
                    else continue;

                    if (string.IsNullOrEmpty(content)) continue;

                    MsTextCategory? msCat = null;

                    var mainMatch = MainLayerRx.Match(content);
                    if (mainMatch.Success)
                    {
                        string which = mainMatch.Groups[1].Value.ToUpperInvariant();
                        msCat = which == "BOTTOM" ? MsTextCategory.MainBottom : MsTextCategory.MainTop;
                    }
                    else if (SectionRx.IsMatch(content)) msCat = MsTextCategory.Section;
                    else if (PhRx.IsMatch(content))      msCat = MsTextCategory.Ph;
                    else if (DetailRx.IsMatch(content))  msCat = MsTextCategory.Detail;

                    if (msCat.HasValue)
                        result.Add((msCat.Value, x, y));
                }
                tr.Commit();
            }
            return result;
        }

        private static Dictionary<string, string> ExtractAutoTitle3(Database db)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var msTexts = ScanModelSpaceTexts(db);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    var vpExtents = new List<(double xMin, double yMin, double xMax, double yMax)>();
                    int vpIdx = 0;
                    foreach (ObjectId id in btr)
                    {
                        var vp = tr.GetObject(id, OpenMode.ForRead) as Viewport;
                        if (vp == null) continue;
                        vpIdx++;
                        if (vpIdx == 1) continue; // pierwszy VP = paperspace overview

                        double cx = vp.ViewCenter.X;
                        double cy = vp.ViewCenter.Y;
                        double h  = vp.ViewHeight;
                        double w  = h * (vp.Width / vp.Height);
                        vpExtents.Add((cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2));
                    }

                    string mainLayer  = null;
                    bool hasSections  = false;
                    bool hasPh        = false;
                    bool hasDetail    = false;

                    foreach (var (cat, x, y) in msTexts)
                    {
                        bool visible = vpExtents.Any(e => x >= e.xMin && x <= e.xMax && y >= e.yMin && y <= e.yMax);
                        if (!visible) continue;

                        switch (cat)
                        {
                            case MsTextCategory.MainBottom:
                                if (mainLayer == null) mainLayer = "BOTTOM LAYER"; break;
                            case MsTextCategory.MainTop:
                                if (mainLayer == null) mainLayer = "TOP LAYER"; break;
                            case MsTextCategory.Section: hasSections = true; break;
                            case MsTextCategory.Ph:      hasPh = true;       break;
                            case MsTextCategory.Detail:  hasDetail = true;   break;
                        }
                    }

                    // kolejność: main_layer, DETAIL, SECTIONS, PH DETAILS
                    var parts = new List<string>();
                    if (mainLayer != null) parts.Add(mainLayer);
                    if (hasDetail)   parts.Add("DETAIL");
                    if (hasSections) parts.Add("SECTIONS");
                    if (hasPh)       parts.Add("PH DETAILS");

                    // Oxford-style: 0→"", 1→"X.", 2→"X & Y.", 3+→"A, B & C."
                    string title3;
                    if (parts.Count == 0)
                        title3 = "";
                    else if (parts.Count == 1)
                        title3 = parts[0] + ".";
                    else if (parts.Count == 2)
                        title3 = parts[0] + " & " + parts[1] + ".";
                    else
                        title3 = string.Join(", ", parts.Take(parts.Count - 1)) + " & " + parts.Last() + ".";

                    result[layout.LayoutName] = title3;
                }
                tr.Commit();
            }

            // Post-processing: identyczne TITLE_3 → suffix "i/N" przed kropką
            var groupsToSuffix = result
                .Where(kv => !string.IsNullOrEmpty(kv.Value))
                .GroupBy(kv => kv.Value, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .ToList();

            foreach (var group in groupsToSuffix)
            {
                string baseTitle = group.Key;
                string baseNoDot = baseTitle.EndsWith(".")
                    ? baseTitle.Substring(0, baseTitle.Length - 1)
                    : baseTitle;

                var sortedLayouts = group
                    .Select(kv => kv.Key)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                int n = sortedLayouts.Count;
                for (int i = 0; i < n; i++)
                    result[sortedLayouts[i]] = $"{baseNoDot} {i + 1}/{n}.";
            }

            return result;
        }

        private static Dictionary<string, string> ExtractLayoutScales(Database db)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    var scales = new HashSet<int>();
                    int vpIdx = 0;
                    foreach (ObjectId id in btr)
                    {
                        var vp = tr.GetObject(id, OpenMode.ForRead) as Viewport;
                        if (vp == null) continue;
                        vpIdx++;
                        if (vpIdx == 1) continue; // paperspace overview

                        if (vp.ViewHeight <= 0) continue;
                        double scale = vp.Height / vp.ViewHeight;
                        if (scale <= 0) continue;
                        int N = (int)Math.Round(1.0 / scale);
                        if (N <= 0) continue;
                        scales.Add(N);
                    }

                    if (scales.Count == 0) { result[layout.LayoutName] = ""; continue; }

                    var sorted = scales.OrderByDescending(n => n).Select(n => "1:" + n).ToList();

                    string fmt;
                    if (sorted.Count == 1)      fmt = sorted[0];
                    else if (sorted.Count == 2) fmt = sorted[0] + " & " + sorted[1];
                    else                        fmt = string.Join(", ", sorted.Take(sorted.Count - 1)) + " & " + sorted.Last();

                    result[layout.LayoutName] = fmt + " @ A1";
                }
                tr.Commit();
            }
            return result;
        }

        private static int RenameLayoutsFromDrawingNumber(Database db)
        {
            var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;

            var planned = new List<(string oldName, string newName)>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string drawingNo = null;
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    foreach (ObjectId oid in btr)
                    {
                        var br = tr.GetObject(oid, OpenMode.ForRead) as BlockReference;
                        if (br == null) continue;
                        if (!string.Equals(br.Name, TitleBlockName, StringComparison.OrdinalIgnoreCase))
                            continue;
                        foreach (ObjectId aid in br.AttributeCollection)
                        {
                            var att = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                            if (att == null) continue;
                            if (string.Equals(att.Tag, "DRAWING_NUMBER", StringComparison.OrdinalIgnoreCase))
                            {
                                drawingNo = att.TextString;
                                break;
                            }
                        }
                        if (drawingNo != null) break;
                    }

                    if (string.IsNullOrEmpty(drawingNo)) continue;

                    int dashIdx = drawingNo.LastIndexOf('-');
                    string suffix  = dashIdx >= 0 ? drawingNo.Substring(dashIdx + 1) : drawingNo;
                    string newName = suffix + "C1";

                    if (string.Equals(layout.LayoutName, newName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    planned.Add((layout.LayoutName, newName));
                }
                tr.Commit();
            }

            if (planned.Count == 0) return 0;

            var lm      = LayoutManager.Current;
            var tempMap = new List<(string tempName, string targetName)>();

            foreach (var op in planned)
            {
                string tempName = "_TMP_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                try
                {
                    lm.RenameLayout(op.oldName, tempName);
                    tempMap.Add((tempName, op.newName));
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nGAI rename phase1 fail ({op.oldName} -> {tempName}): {ex.Message}");
                }
            }

            int renamed = 0;
            foreach (var op in tempMap)
            {
                try
                {
                    lm.RenameLayout(op.tempName, op.targetName);
                    renamed++;
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nGAI rename phase2 fail ({op.tempName} -> {op.targetName}): {ex.Message}");
                }
            }

            return renamed;
        }

        private static int? ExtractFirstGaDrawingNumber(Database gaDb)
        {
            var layoutNames = new List<string>();
            var drawingNumbers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using (var tr = gaDb.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(gaDb.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string drawingNo = null;
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    foreach (ObjectId oid in btr)
                    {
                        var br = tr.GetObject(oid, OpenMode.ForRead) as BlockReference;
                        if (br == null) continue;
                        if (!string.Equals(br.Name, TitleBlockName, StringComparison.OrdinalIgnoreCase))
                            continue;
                        foreach (ObjectId aid in br.AttributeCollection)
                        {
                            var att = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                            if (att == null) continue;
                            if (string.Equals(att.Tag, "DRAWING_NUMBER", StringComparison.OrdinalIgnoreCase))
                            {
                                drawingNo = att.TextString;
                                break;
                            }
                        }
                        break;
                    }

                    if (!string.IsNullOrEmpty(drawingNo))
                    {
                        layoutNames.Add(layout.LayoutName);
                        drawingNumbers[layout.LayoutName] = drawingNo;
                    }
                }
                tr.Commit();
            }

            if (layoutNames.Count == 0) return null;

            layoutNames.Sort(StringComparer.OrdinalIgnoreCase);
            string firstDrawingNo = drawingNumbers[layoutNames[0]];

            int dashIdx = firstDrawingNo.LastIndexOf('-');
            if (dashIdx < 0 || dashIdx == firstDrawingNo.Length - 1) return null;
            string suffix = firstDrawingNo.Substring(dashIdx + 1).Trim();

            var m = GaSuffixNumberRx.Match(suffix);
            if (!m.Success) return null;

            if (int.TryParse(m.Groups[1].Value, out int num)) return num;
            return null;
        }

        private static string BuildDrawingSuffix(int? plotNumber, int layoutIdx0Based, int? firstGaNumber = null)
        {
            if (firstGaNumber.HasValue)
            {
                int num = firstGaNumber.Value + layoutIdx0Based;
                return "RC" + num.ToString();
            }
            if (plotNumber.HasValue)
            {
                int p = plotNumber.Value;
                if (p < 10)
                    return "RC" + p.ToString("D2") + layoutIdx0Based.ToString();
                else
                    return "RC" + p.ToString() + (layoutIdx0Based + 1).ToString();
            }
            return "RC" + (layoutIdx0Based + 1).ToString("D3");
        }

        private static Dictionary<string, string> ExtractSlabValuesFromDb(Database db, string layerName)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    foreach (ObjectId id in btr)
                    {
                        var mt = tr.GetObject(id, OpenMode.ForRead) as MText;
                        if (mt == null) continue;

                        string contents = mt.Contents;
                        if (!contents.Contains("SLAB AREA")) continue;

                        var m1 = SlabAreaExtractRx.Match(contents);
                        if (m1.Success)
                        {
                            result[KeySlabArea] = m1.Groups[1].Value;
                        }
                        else
                        {
                            var rm = SlabAreaRawExtractRx.Match(contents);
                            if (rm.Success) { string raw = rm.Groups[1].Value.Trim(); if (!string.IsNullOrEmpty(raw)) result[KeySlabAreaRaw] = raw; }
                        }

                        var m2 = SlabPerimeterExtractRx.Match(contents);
                        if (m2.Success)
                        {
                            result[KeySlabPerimeter] = m2.Groups[1].Value;
                        }
                        else
                        {
                            var rm = SlabPerimeterRawExtractRx.Match(contents);
                            if (rm.Success) { string raw = rm.Groups[1].Value.Trim(); if (!string.IsNullOrEmpty(raw)) result[KeySlabPerimeterRaw] = raw; }
                        }

                        var m3 = SlabThicknessExtractRx.Match(contents);
                        if (m3.Success)
                        {
                            result[KeySlabThickness] = m3.Groups[1].Value;
                        }
                        else
                        {
                            var rm = SlabThicknessRawExtractRx.Match(contents);
                            if (rm.Success) { string raw = rm.Groups[1].Value.Trim(); if (!string.IsNullOrEmpty(raw)) result[KeySlabThicknessRaw] = raw; }
                        }

                        var vMatch = ConcreteVolumeExtractRx.Match(contents);
                        if (vMatch.Success)
                        {
                            result[KeyConcreteVolume] = vMatch.Groups[2].Value;
                        }
                        else
                        {
                            var rm = ConcreteVolumeRawExtractRx.Match(contents);
                            if (rm.Success) { string raw = rm.Groups[1].Value.Trim(); if (!string.IsNullOrEmpty(raw)) result[KeyConcreteVolumeRaw] = raw; }
                        }

                        var dMatch = ConcreteDesignatedBlockRx.Match(contents);
                        if (dMatch.Success)
                            result[KeyConcreteDesignated] = dMatch.Value;

                        tr.Commit();
                        return result;
                    }
                }
                tr.Commit();
            }
            return result;
        }

        private static int ApplySlabValuesToActiveDb(
            Database db,
            Dictionary<string, string> values,
            string layerName)
        {
            int updatedLayouts = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    bool layoutTouched = false;

                    foreach (ObjectId id in btr)
                    {
                        var mt = tr.GetObject(id, OpenMode.ForRead) as MText;
                        if (mt == null) continue;

                        string contents = mt.Contents;
                        if (!contents.Contains("SLAB AREA")) continue;

                        string newContents = contents;

                        string vArea    = values.TryGetValue(KeySlabArea,          out var a)   && !string.IsNullOrEmpty(a)   ? a   : null;
                        string vAreaRaw = values.TryGetValue(KeySlabAreaRaw,       out var ar)  && !string.IsNullOrEmpty(ar)  ? ar  : null;
                        string vPer     = values.TryGetValue(KeySlabPerimeter,     out var per) && !string.IsNullOrEmpty(per) ? per : null;
                        string vPerRaw  = values.TryGetValue(KeySlabPerimeterRaw,  out var prr) && !string.IsNullOrEmpty(prr) ? prr : null;
                        string vTh      = values.TryGetValue(KeySlabThickness,     out var th)  && !string.IsNullOrEmpty(th)  ? th  : null;
                        string vThRaw   = values.TryGetValue(KeySlabThicknessRaw,  out var thr) && !string.IsNullOrEmpty(thr) ? thr : null;
                        string vVol     = values.TryGetValue(KeyConcreteVolume,    out var vol) && !string.IsNullOrEmpty(vol) ? vol : null;
                        string vVolRaw  = values.TryGetValue(KeyConcreteVolumeRaw, out var vr)  && !string.IsNullOrEmpty(vr)  ? vr  : null;

                        if (vArea != null)
                        {
                            string before = newContents;
                            newContents = SlabAreaReplaceRx.Replace(newContents, "${1}" + vArea + "${3}");
                            if (newContents == before)
                                newContents = SlabAreaUniversalRx.Replace(newContents, "${1}" + vArea + " m");
                        }
                        else if (vAreaRaw != null)
                            newContents = SlabAreaRawReplaceRx.Replace(newContents, "${1}" + vAreaRaw);
                        else
                            newContents = SlabAreaReplaceRx.Replace(newContents, "${1}");

                        if (vPer != null)
                        {
                            string before = newContents;
                            newContents = SlabPerimeterReplaceRx.Replace(newContents, "${1}" + vPer + "${3}");
                            if (newContents == before)
                                newContents = SlabPerimeterUniversalRx.Replace(newContents, "${1}" + vPer + " m");
                        }
                        else if (vPerRaw != null)
                            newContents = SlabPerimeterRawReplaceRx.Replace(newContents, "${1}" + vPerRaw);
                        else
                            newContents = SlabPerimeterReplaceRx.Replace(newContents, "${1}");

                        if (vTh != null)
                        {
                            string before = newContents;
                            newContents = SlabThicknessReplaceRx.Replace(newContents, "${1}" + vTh + "${3}");
                            if (newContents == before)
                                newContents = SlabThicknessUniversalRx.Replace(newContents, "${1}" + vTh + " mm");
                        }
                        else if (vThRaw != null)
                            newContents = SlabThicknessRawReplaceRx.Replace(newContents, "${1}" + vThRaw);
                        else
                            newContents = SlabThicknessReplaceRx.Replace(newContents, "${1}");

                        // CONCRETE VOLUME — podmień tylko liczbę; m³ codes z RC zachowane, tail wycinany
                        if (vVol != null)
                        {
                            string before = newContents;
                            newContents = ConcreteVolumeReplaceRx.Replace(newContents, m => m.Groups[1].Value + vVol + m.Groups[3].Value);
                            if (newContents == before)
                                newContents = ConcreteVolumeUniversalRx.Replace(newContents, "${1}" + vVol + " m³");
                        }
                        else if (vVolRaw != null)
                            newContents = ConcreteVolumeRawReplaceRx.Replace(newContents, "${1}" + vVolRaw);
                        else
                            newContents = ConcreteVolumeReplaceRx.Replace(newContents, m => m.Groups[1].Value + m.Groups[3].Value);

                        // CONCRETE TO BE DESIGNATED block — substring-based (unika interpretacji $ w gaBlock)
                        if (values.TryGetValue(KeyConcreteDesignated, out var gaBlock) && !string.IsNullOrEmpty(gaBlock))
                        {
                            var rcMatch = ConcreteDesignatedBlockRx.Match(newContents);
                            if (rcMatch.Success)
                            {
                                newContents = newContents.Substring(0, rcMatch.Index)
                                            + gaBlock
                                            + newContents.Substring(rcMatch.Index + rcMatch.Length);
                            }
                        }

                        // HYSTOOLS — podmień DK wariant wg thickness (225→DK90, 300→DK165)
                        if (values.TryGetValue(KeySlabThickness, out var thStr) &&
                            int.TryParse(thStr, out int thMm))
                        {
                            string targetDk = thMm == 225 ? "DK90"
                                           : thMm == 300 ? "DK165"
                                           : null;
                            if (targetDk != null)
                                newContents = HystoolsRx.Replace(newContents, "HYSTOOLS " + targetDk);
                        }

                        if (!string.Equals(newContents, contents, StringComparison.Ordinal))
                        {
                            mt.UpgradeOpen();
                            mt.Contents = newContents;
                            layoutTouched = true;
                        }
                    }

                    if (layoutTouched) updatedLayouts++;
                }
                tr.Commit();
            }
            return updatedLayouts;
        }

        private static Dictionary<string, string> ReadA1BLAttributesFromFile(string path)
        {
            using (var sideDb = new Database(false, true))
            {
                bool isDxf = path.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase);
                if (isDxf)
                    sideDb.DxfIn(path, null);
                else
                    sideDb.ReadDwgFile(path, System.IO.FileShare.Read, true, "");

                return ExtractA1BLAttributes(sideDb);
            }
        }

        private static Dictionary<string, string> ExtractA1BLAttributes(Database db)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    foreach (ObjectId id in btr)
                    {
                        var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (br == null) continue;
                        if (!string.Equals(br.Name, TitleBlockName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        foreach (ObjectId attId in br.AttributeCollection)
                        {
                            var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                            if (att == null) continue;
                            if (!result.ContainsKey(att.Tag))
                                result[att.Tag] = att.TextString;
                        }

                        if (result.Count > 0)
                        {
                            tr.Commit();
                            return result;
                        }
                    }
                }
                tr.Commit();
            }
            return result;
        }

        private static int ApplyA1BLAttributesToActiveDb(
            Database db,
            Dictionary<string, string> src,
            string[] tagsToCopy,
            string gaTitlePrefix,
            string gaDrawingPrefix,
            int? plotNumberFromGa,
            int? firstGaNumber = null)
        {
            int updatedLayouts = 0;
            var tagsSet = new HashSet<string>(tagsToCopy, StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                // Posortuj alfabetycznie — określa indeks każdego layoutu dla DRAWING_NUMBER
                var sortedLayoutNames = new List<string>();
                foreach (DBDictionaryEntry sortEntry in layoutDict)
                {
                    var sortLayout = tr.GetObject(sortEntry.Value, OpenMode.ForRead) as Layout;
                    if (sortLayout == null) continue;
                    if (string.Equals(sortLayout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase)) continue;
                    sortedLayoutNames.Add(sortLayout.LayoutName);
                }
                sortedLayoutNames.Sort(StringComparer.OrdinalIgnoreCase);
                var layoutNameToIdx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < sortedLayoutNames.Count; i++)
                    layoutNameToIdx[sortedLayoutNames[i]] = i;

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    int currentLayoutIdx = layoutNameToIdx.ContainsKey(layout.LayoutName)
                        ? layoutNameToIdx[layout.LayoutName] : 0;
                    int? currentLayoutPlot = plotNumberFromGa;

                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    bool layoutTouched = false;

                    foreach (ObjectId id in btr)
                    {
                        var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (br == null) continue;
                        if (!string.Equals(br.Name, TitleBlockName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        foreach (ObjectId attId in br.AttributeCollection)
                        {
                            var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                            if (att == null) continue;

                            string newVal = null;

                            // Specjalna obsługa TITLE_1: podmień lub usuń prefix przed "REINFORCEMENT DETAILS"
                            if (string.Equals(att.Tag, "TITLE_1", StringComparison.OrdinalIgnoreCase))
                            {
                                newVal = ReplaceRcTitlePrefix(att.TextString, gaTitlePrefix ?? "");
                            }
                            // DRAWING_NUMBER: buduj suffix wg numeru plotu i indeksu layoutu
                            else if (!string.IsNullOrEmpty(gaDrawingPrefix) &&
                                     string.Equals(att.Tag, "DRAWING_NUMBER", StringComparison.OrdinalIgnoreCase))
                            {
                                string suffix = BuildDrawingSuffix(currentLayoutPlot, currentLayoutIdx, firstGaNumber);
                                newVal = gaDrawingPrefix + "-" + suffix;
                            }
                            // Standardowa obsługa CLIENT_* / PROJ_*
                            else if (tagsSet.Contains(att.Tag))
                            {
                                if (src.TryGetValue(att.Tag, out string srcVal))
                                    newVal = !string.IsNullOrEmpty(srcVal) ? srcVal : "";
                            }

                            if (newVal == null) continue;
                            if (string.Equals(att.TextString, newVal, StringComparison.Ordinal)) continue;

                            att.UpgradeOpen();
                            att.TextString = newVal;
                            layoutTouched = true;
                        }
                    }

                    if (layoutTouched) updatedLayouts++;
                }
                tr.Commit();
            }
            return updatedLayouts;
        }

        private static int ApplyAutoFieldsToActiveDb(
            Database db,
            Dictionary<string, string> autoTitle3Map,
            Dictionary<string, string> autoScalesMap,
            string nowDate)
        {
            int updatedLayouts = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    bool layoutChanged = false;

                    foreach (ObjectId id in btr)
                    {
                        var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (br == null) continue;
                        if (!string.Equals(br.Name, TitleBlockName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        foreach (ObjectId attId in br.AttributeCollection)
                        {
                            var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                            if (att == null) continue;

                            string newVal = null;

                            if (autoTitle3Map != null &&
                                string.Equals(att.Tag, "TITLE_3", StringComparison.OrdinalIgnoreCase))
                            {
                                if (autoTitle3Map.TryGetValue(layout.LayoutName, out string t3)
                                    && !string.IsNullOrEmpty(t3))
                                {
                                    newVal = t3;
                                }
                            }
                            else if (autoScalesMap != null &&
                                     string.Equals(att.Tag, "SCALE", StringComparison.OrdinalIgnoreCase))
                            {
                                if (autoScalesMap.TryGetValue(layout.LayoutName, out string sc)
                                    && !string.IsNullOrEmpty(sc))
                                {
                                    newVal = sc;
                                }
                            }
                            else if (!string.IsNullOrEmpty(nowDate) &&
                                     string.Equals(att.Tag, "DATE", StringComparison.OrdinalIgnoreCase))
                            {
                                newVal = nowDate;
                            }

                            if (newVal != null && !string.Equals(att.TextString, newVal, StringComparison.Ordinal))
                            {
                                att.UpgradeOpen();
                                att.TextString = newVal;
                                layoutChanged = true;
                            }
                        }
                    }

                    if (layoutChanged) updatedLayouts++;
                }
                tr.Commit();
            }

            return updatedLayouts;
        }

        // Wstawia ramki z odnośnikami "SEE DRG ..." na każdym layoutcie RC.
        // Każdy layout dostaje ramki do wszystkich POZOSTAŁYCH layoutów.
        // Wymaga że TITLE_3 i DRAWING_NUMBER są już zapisane w atrybutach A1-BL.
        // Zwraca łączną liczbę wstawionych ramek.
        private static int CreateReferenceFramesOnAllLayouts(Database db)
        {
            var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;

            // 1. Zbierz dane: layoutName → (title3, drawingNo)
            var layoutData = new List<(string name, string title3, string drawingNo)>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string title3 = null, drawingNo = null;
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    foreach (ObjectId oid in btr)
                    {
                        var br = tr.GetObject(oid, OpenMode.ForRead) as BlockReference;
                        if (br == null) continue;
                        if (!string.Equals(br.Name, TitleBlockName, StringComparison.OrdinalIgnoreCase))
                            continue;
                        foreach (ObjectId aid in br.AttributeCollection)
                        {
                            var att = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                            if (att == null) continue;
                            if (string.Equals(att.Tag, "TITLE_3", StringComparison.OrdinalIgnoreCase))
                                title3 = att.TextString;
                            else if (string.Equals(att.Tag, "DRAWING_NUMBER", StringComparison.OrdinalIgnoreCase))
                                drawingNo = att.TextString;
                        }
                        break;
                    }

                    if (!string.IsNullOrEmpty(title3) && !string.IsNullOrEmpty(drawingNo))
                        layoutData.Add((layout.LayoutName, title3, drawingNo));
                }
                tr.Commit();
            }

            layoutData = layoutData.OrderBy(d => d.name, StringComparer.OrdinalIgnoreCase).ToList();

            if (layoutData.Count < 2)
            {
                ed.WriteMessage($"\nRCN-REFS: not enough layouts ({layoutData.Count}), no frames to insert.");
                return 0;
            }

            // 2. Upewnij się że warstwa ASD-RCN-REFS istnieje
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(RcnRefsLayer))
                {
                    lt.UpgradeOpen();
                    var ltr = new LayerTableRecord
                    {
                        Name = RcnRefsLayer,
                        Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                            Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 7)
                    };
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
                else
                {
                    var ltr = (LayerTableRecord)tr.GetObject(lt[RcnRefsLayer], OpenMode.ForWrite);
                    ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                        Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 7);
                }
                tr.Commit();
            }

            // 3. Dla każdego layoutu: wyczyść warstwę i wstaw ramki
            const double frameWidth  = 95;
            const double frameHeight = 15;
            const double rowGap      = 4;
            const double startX      = 900;
            const double startY      = 580;
            const double innerPad    = 2;

            int totalFrames = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var styleTable = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                ObjectId textStyleId = ObjectId.Null;
                if (styleTable.Has("ROMANS NARROW"))
                    textStyleId = styleTable["ROMANS NARROW"];

                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                    // 3a. Wyczyść poprzednie ramki na ASD-RCN-REFS
                    var toDelete = new List<ObjectId>();
                    foreach (ObjectId oid in btr)
                    {
                        var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;
                        if (string.Equals(ent.Layer, RcnRefsLayer, StringComparison.OrdinalIgnoreCase))
                            toDelete.Add(oid);
                    }
                    foreach (var id in toDelete)
                    {
                        var ent = (Entity)tr.GetObject(id, OpenMode.ForWrite);
                        ent.Erase();
                    }

                    // 3b. Wstaw ramki dla pozostałych layoutów
                    int frameIdx = 0;
                    foreach (var other in layoutData)
                    {
                        if (string.Equals(other.name, layout.LayoutName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        double y = startY - (frameIdx * (frameHeight + rowGap));

                        var pl = new Polyline();
                        pl.AddVertexAt(0, new Point2d(startX, y), 0, 0, 0);
                        pl.AddVertexAt(1, new Point2d(startX + frameWidth, y), 0, 0, 0);
                        pl.AddVertexAt(2, new Point2d(startX + frameWidth, y - frameHeight), 0, 0, 0);
                        pl.AddVertexAt(3, new Point2d(startX, y - frameHeight), 0, 0, 0);
                        pl.Closed = true;
                        pl.Layer = RcnRefsLayer;
                        pl.ColorIndex = 7;
                        btr.AppendEntity(pl);
                        tr.AddNewlyCreatedDBObject(pl, true);

                        string titleNoDot = (other.title3 ?? "").TrimEnd('.', ' ');
                        var mt = new MText();
                        mt.Location = new Point3d(startX + innerPad, y - innerPad, 0);
                        mt.Width = frameWidth - 2 * innerPad;
                        mt.TextHeight = 4.0;
                        mt.Attachment = AttachmentPoint.TopLeft;
                        mt.Layer = RcnRefsLayer;
                        if (!textStyleId.IsNull) mt.TextStyleId = textStyleId;
                        mt.Contents = $"\\C3;{titleNoDot}\\PSEE DRG. {other.drawingNo}";
                        btr.AppendEntity(mt);
                        tr.AddNewlyCreatedDBObject(mt, true);

                        frameIdx++;
                        totalFrames++;
                    }
                }
                tr.Commit();
            }

            return totalFrames;
        }

        // ── PANEL 4: QA VALIDATOR ─────────────────────────────────────────────

        [CommandMethod("ASD-BBSV")]
        public void CmdSprawdzBBS()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Sprawdz BBS — R87/R95/R81/R83/R92\n");
        }

        [CommandMethod("ASD-PIV")]
        public void CmdPIVCheck()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: PIV Check 15R — centralny validator\n");
        }

        [CommandMethod("ASD-GER")]
        public void CmdRaportBledow()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Raport Bledow — PIV_Dashboard.xlsx\n");
        }

        [CommandMethod("ASD-QAP")]
        public void CmdPodgladQA()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Podglad QA — live status regul\n");
        }

        // ── PANEL 5: EKSPORT ──────────────────────────────────────────────────

        [CommandMethod("ASD-BSX")]
        public void CmdGenerujBBS()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Generuj BBS — BS8666 Excel z waga\n");
        }

        [CommandMethod("ASD-PDF")]
        public void CmdPDFExport()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: PDF Export — dostepne w Sprint 2\n");
        }

        [CommandMethod("ASD-CAG")]
        public void CmdCalcDoc()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Calc Doc — Template_PiledRaft.docx\n");
        }

        [CommandMethod("ASD-TRX")]
        public void CmdTransmittal()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Transmittal — lista PDF do wyslania\n");
        }
    }
}
