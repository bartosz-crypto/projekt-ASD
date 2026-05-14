using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
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
        private const string KeySlabArea             = "SLAB_AREA";
        private const string KeySlabPerimeter        = "SLAB_PERIMETER";
        private const string KeySlabThickness        = "SLAB_THICKNESS";
        private const string KeyConcreteVolume       = "CONCRETE_VOLUME";
        private const string KeyConcreteVolumeTail   = "CONCRETE_VOLUME_TAIL";
        private const string KeyConcreteDesignated   = "CONCRETE_DESIGNATED_BLOCK";

        private static readonly string[] GaiFieldsToCopy = new[]
        {
            "CLIENT_1", "CLIENT_2", "CLIENT_3",
            "PROJ_1",   "PROJ_2",   "PROJ_3"
        };

        // Slab extract: group 1 = numeric value
        private static readonly Regex SlabAreaExtractRx      = new Regex(@"SLAB\s+AREA\s*=\s*([\d.]+)\s*m",       RegexOptions.IgnoreCase);
        private static readonly Regex SlabPerimeterExtractRx = new Regex(@"SLAB\s+PERIMETER\s*=\s*([\d.]+)\s*m",  RegexOptions.IgnoreCase);
        private static readonly Regex SlabThicknessExtractRx = new Regex(@"SLAB\s+THICKNESS\s*=\s*([\d.]+)\s*mm", RegexOptions.IgnoreCase);

        // CONCRETE VOLUME — g1=prefix, g2=number, g3=format codes (\H..\S3..\H..), g4=tail
        private static readonly Regex ConcreteVolumeExtractRx = new Regex(
            @"(CONCRETE\s+VOLUME\s*=\s*)([\d.]+)(\s*m\\H[^;]*;\\S3\^\s*;\\H[^;]*;)([^\\]*)",
            RegexOptions.IgnoreCase);

        // Slab replace: group 1 = "X = ", group 2 = number, group 3 = " m"/" mm"
        private static readonly Regex SlabAreaReplaceRx      = new Regex(@"(SLAB\s+AREA\s*=\s*)([\d.]+)(\s*m)",       RegexOptions.IgnoreCase);
        private static readonly Regex SlabPerimeterReplaceRx = new Regex(@"(SLAB\s+PERIMETER\s*=\s*)([\d.]+)(\s*m)",  RegexOptions.IgnoreCase);
        private static readonly Regex SlabThicknessReplaceRx = new Regex(@"(SLAB\s+THICKNESS\s*=\s*)([\d.]+)(\s*mm)", RegexOptions.IgnoreCase);

        // CONCRETE TO BE DESIGNATED block — od frazy do "CERTIFICATE." (włącznie z kropką)
        private static readonly Regex ConcreteDesignatedBlockRx = new Regex(
            @"CONCRETE\s+TO\s+BE\s+DESIGNATED.*?CERTIFICATE\s*\.",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        // Wyciąga numer PLOT z TITLE_1 (np. "PLOT 3" → 3)
        private static readonly Regex TitlePlotNumberRx =
            new Regex(@"\bPLOT\s+(\d+)", RegexOptions.IgnoreCase);

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
                Title  = "Wczytaj projekt",
                Filter = "Project JSON (*.json)|*.json",
                DefaultExt = ".json"
            };

            if (dlg.ShowDialog() != true) return;

            try
            {
                string json = File.ReadAllText(dlg.FileName);
                SessionData.CurrentProject = JsonConvert.DeserializeObject<ProjectData>(json);
                doc.Editor.WriteMessage(
                    $"\nWczytano: {SessionData.CurrentProject.ProjectName}" +
                    $", DRW: {SessionData.CurrentProject.DRWNumber}" +
                    $", h={SessionData.CurrentProject.SlabThickness}mm\n");
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\nBlad wczytywania projektu: {ex.Message}\n");
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
                Title  = "Wybierz plik GA (źródło tabelki tytułowej)",
                Filter = "AutoCAD drawings (*.dwg;*.dxf)|*.dwg;*.dxf"
            };
            if (dlg.ShowDialog() != true) return;
            string path = dlg.FileName;
            ed.WriteMessage($"\nGAI: czytam GA z '{Path.GetFileName(path)}'...");

            // 2. Czytaj atrybuty A1-BL (zawiera też TITLE_1)
            Dictionary<string, string> srcAttrs;
            try
            {
                srcAttrs = ReadA1BLAttributesFromFile(path);
            }
            catch (System.Exception ex)
            {
                string inner = ex.InnerException != null
                    ? $"\nInnerException: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}"
                    : "";
                string msg = $"Nie można otworzyć pliku GA:\n" +
                             $"{ex.GetType().Name}: {ex.Message}{inner}\n\n" +
                             $"Stack (top 5):\n{string.Join("\n", (ex.StackTrace ?? "").Split('\n').Take(5))}";
                MessageBox.Show(msg, "GAI", MessageBoxButton.OK, MessageBoxImage.Error);
                ed.WriteMessage($"\nGAI EXCEPTION: {ex}");
                return;
            }

            if (srcAttrs == null || srcAttrs.Count == 0)
            {
                MessageBox.Show($"W pliku GA nie znaleziono bloku '{TitleBlockName}'.",
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

            // 4. Otwórz GA drugi raz dla SLAB values (MText na PCN-Text)
            var gaSlabValues = new Dictionary<string, string>();
            try
            {
                using (var sideDb = new Database(false, true))
                {
                    bool isDxf = path.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase);
                    if (isDxf) sideDb.DxfIn(path, null);
                    else       sideDb.ReadDwgFile(path, System.IO.FileShare.Read, true, "");

                    gaSlabValues = ExtractSlabValuesFromDb(sideDb, GaSlabNotesLayer);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Błąd przy czytaniu GA (SLAB notes):\n{ex.Message}",
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

            // 5b. Auto-wykryj TITLE_3 z Model space + viewportów RC
            Dictionary<string, string> autoTitle3Map;
            try
            {
                autoTitle3Map = ExtractAutoTitle3(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nGAI: nie udało się auto-wykryć TITLE_3: {ex.Message}");
                autoTitle3Map = new Dictionary<string, string>();
            }

            // 5c. SCALE — auto z viewportów per layout
            Dictionary<string, string> autoScalesMap;
            try
            {
                autoScalesMap = ExtractLayoutScales(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nGAI: nie udało się wykryć skali viewportów: {ex.Message}");
                autoScalesMap = new Dictionary<string, string>();
            }

            // 5d. DATE — aktualny miesiąc + rok
            string nowDate = DateTime.Now.ToString(
                "MMM yyyy",
                System.Globalization.CultureInfo.InvariantCulture).ToUpper();

            // 6. Preview + confirm
            var sb = new StringBuilder();
            sb.AppendLine("Skopiować poniższe wartości do RC?");
            sb.AppendLine();
            sb.AppendLine("DANE PROJEKTU:");
            foreach (var tag in GaiFieldsToCopy)
            {
                string val = srcAttrs.TryGetValue(tag, out var v) ? v : "(brak)";
                sb.AppendLine($"  {tag,-10}: {val}");
            }
            sb.AppendLine();
            sb.AppendLine("TITLE PREFIX (przed REINFORCEMENT DETAILS):");
            sb.AppendLine($"  {gaTitlePrefix ?? "(brak)"}");
            sb.AppendLine();
            sb.AppendLine("DRAWING_NUMBER:");
            if (string.IsNullOrEmpty(gaDrawingPrefix))
            {
                sb.AppendLine("  (brak prefiksu z GA — nie nadpiszemy)");
            }
            else
            {
                string plotStr = plotNumberFromGa.HasValue
                    ? $"PLOT {plotNumberFromGa.Value} (z GA prefix)"
                    : "brak PLOT N w GA prefix";
                sb.AppendLine($"  Plot: {plotStr}");

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
                    string suffix = BuildDrawingSuffix(plotNumberFromGa, i);
                    sb.AppendLine($"  {rcLayoutNames[i]} → {gaDrawingPrefix}-{suffix}");
                }
            }
            sb.AppendLine();
            sb.AppendLine("TITLE_3 (auto-wykrycie z Model space + viewporty):");
            if (autoTitle3Map == null || autoTitle3Map.Count == 0)
            {
                sb.AppendLine("  (brak — automat nic nie wykrył)");
            }
            else
            {
                var sorted = autoTitle3Map.OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase).ToList();
                foreach (var kv in sorted)
                {
                    string val = string.IsNullOrEmpty(kv.Value)
                        ? "(brak wykrycia — zostawiam istniejący)"
                        : "\"" + kv.Value + "\"";
                    sb.AppendLine($"  {kv.Key} → {val}");
                }
            }
            sb.AppendLine();
            sb.AppendLine("SCALE (auto z viewportów):");
            if (autoScalesMap == null || autoScalesMap.Count == 0)
            {
                sb.AppendLine("  (brak)");
            }
            else
            {
                foreach (var kv in autoScalesMap.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
                {
                    string val = string.IsNullOrEmpty(kv.Value) ? "(brak viewportów)" : "\"" + kv.Value + "\"";
                    sb.AppendLine($"  {kv.Key} → {val}");
                }
            }
            sb.AppendLine();
            sb.AppendLine($"DATE: {nowDate}");
            sb.AppendLine();
            sb.AppendLine("LAYOUT RENAME (po nadpisaniu DRAWING_NUMBER):");
            sb.AppendLine("  (nowe nazwy = suffix po '-' + 'C1', np. RC030C1)");
            sb.AppendLine();
            sb.AppendLine("SLAB NOTES (wartości liczbowe):");
            sb.AppendLine($"  SLAB AREA       : {(gaSlabValues.TryGetValue(KeySlabArea,       out var sa)  ? sa  : "(brak)")} m²");
            sb.AppendLine($"  SLAB PERIMETER  : {(gaSlabValues.TryGetValue(KeySlabPerimeter,  out var spe) ? spe : "(brak)")} m");
            sb.AppendLine($"  SLAB THICKNESS  : {(gaSlabValues.TryGetValue(KeySlabThickness,  out var sth) ? sth : "(brak)")} mm");
            sb.AppendLine($"  CONCRETE VOLUME : {(gaSlabValues.TryGetValue(KeyConcreteVolume, out var svo) ? svo : "(brak)")} m³");
            sb.AppendLine();
            sb.AppendLine("CONCRETE VOLUME tail (po wartości):");
            string volTail = gaSlabValues.TryGetValue(KeyConcreteVolumeTail, out var vt) ? vt : "";
            sb.AppendLine($"  {(string.IsNullOrWhiteSpace(volTail) ? "(pusty — RC straci ewentualne 'inc. PILE CAP')" : volTail)}");
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
                sb.AppendLine("  (brak — nie znaleziono w GA)");
            }

            var confirmResult = MessageBox.Show(sb.ToString(), "GAI — potwierdź",
                                                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirmResult != MessageBoxResult.Yes)
            {
                ed.WriteMessage("\nGAI: anulowano przez użytkownika.");
                return;
            }

            // 6. Apply: atrybuty A1-BL + TITLE_1 prefix
            int updatedAttrLayouts;
            try
            {
                updatedAttrLayouts = ApplyA1BLAttributesToActiveDb(
                    db, srcAttrs, GaiFieldsToCopy,
                    gaTitlePrefix, gaDrawingPrefix, plotNumberFromGa, autoTitle3Map,
                    autoScalesMap, nowDate);
            }
            catch (System.Exception ex)
            {
                string inner = ex.InnerException != null
                    ? $"\nInnerException: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}"
                    : "";
                string msg = $"Błąd podczas nadpisywania atrybutów:\n" +
                             $"{ex.GetType().Name}: {ex.Message}{inner}\n\n" +
                             $"Stack (top 5):\n{string.Join("\n", (ex.StackTrace ?? "").Split('\n').Take(5))}";
                MessageBox.Show(msg, "GAI", MessageBoxButton.OK, MessageBoxImage.Error);
                ed.WriteMessage($"\nGAI EXCEPTION: {ex}");
                return;
            }

            // 7a. Rename layoutów na podstawie nadpisanych DRAWING_NUMBER
            int renamedLayouts = 0;
            try
            {
                renamedLayouts = RenameLayoutsFromDrawingNumber(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nGAI: rename layoutów nie powiódł się: {ex.Message}");
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
                    ed.WriteMessage($"\nGAI: ostrzeżenie SLAB: {ex.Message}");
                }
            }

            // 8. Log + komunikat
            ed.WriteMessage($"\nGAI: zaktualizowano {updatedAttrLayouts} layout(ów) z tabelką, " +
                            $"{updatedSlabLayouts} layout(ów) z SLAB NOTES, " +
                            $"renamowano {renamedLayouts} layout(ów).");
            MessageBox.Show($"Zaktualizowano:\n• tabelka tytułowa: {updatedAttrLayouts} layout(ów)\n" +
                            $"• SLAB NOTES: {updatedSlabLayouts} layout(ów)\n" +
                            $"• Zmieniono nazwy: {renamedLayouts} layout(ów)",
                            "GAI", MessageBoxButton.OK, MessageBoxImage.Information);
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

            var peo = new PromptEntityOptions("\nWskaż obrys płyty (warstwa SD-PILED-RAFT): ");
            peo.SetRejectMessage("\nWybierz encję.");
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            string validationError = null;
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var ent = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                if (ent == null)
                    validationError = "Nie można otworzyć encji.";
                else if (ent.Layer != "SD-PILED-RAFT")
                    validationError = "Zaznacz polilinię na warstwie SD-PILED-RAFT.";
                else if (!(ent is Polyline))
                    validationError = "Zaznaczona encja nie jest polilinią (LWPolyline).";
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
                    "Najpierw utwórz pręty H10 (rbcr_def_bar_bv) i zarejestruj je komendą ASD-GSETUP.",
                    "ASD-GBOT", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            var result = ReinforcementGenerator.GenerateBottomAsd(doc, per.ObjectId,
                SessionData.TemplateBarsB);
            if (!string.IsNullOrEmpty(result.Error))
            {
                ed.WriteMessage($"\nGBOT błąd: {result.Error}\n");
                System.Windows.MessageBox.Show(result.Error, "ASD-GBOT — błąd",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            SessionData.LapPositionsB1 = result.LapPositionsX;
            SessionData.LapPositionsB2 = result.LapPositionsY;
            ed.WriteMessage($"\nB1/B2: wysyłanie {result.BarsDrawn} prętów do ASD...\n");
        }

        [CommandMethod("ASD-GTOP")]
        public void CmdGenerujT1T2()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed  = doc.Editor;

            var peo = new PromptEntityOptions("\nWskaż obrys płyty (warstwa SD-PILED-RAFT): ");
            peo.SetRejectMessage("\nWybierz encję.");
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            string validationError = null;
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var ent = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                if (ent == null)
                    validationError = "Nie można otworzyć encji.";
                else if (ent.Layer != "SD-PILED-RAFT")
                    validationError = "Zaznacz polilinię na warstwie SD-PILED-RAFT.";
                else if (!(ent is Polyline))
                    validationError = "Zaznaczona encja nie jest polilinią (LWPolyline).";
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
                    "Najpierw utwórz pręty H12 (rbcr_def_bar_bv) i zarejestruj je komendą ASD-GSETUP.",
                    "ASD-GTOP", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            var result = ReinforcementGenerator.GenerateTopAsd(doc, per.ObjectId,
                SessionData.TemplateBarsT,
                SessionData.LapPositionsB1, SessionData.LapPositionsB2);
            if (!string.IsNullOrEmpty(result.Error))
            {
                ed.WriteMessage($"\nGTOP błąd: {result.Error}\n");
                System.Windows.MessageBox.Show(result.Error, "ASD-GTOP — błąd",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            ed.WriteMessage($"\nT1/T2: wysyłanie {result.BarsDrawn} prętów do ASD...\n");
        }

        [CommandMethod("ASD-BMM")]
        public void CmdOznaczPrety()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;

            var fileDlg = new Microsoft.Win32.OpenFileDialog
            {
                Title      = "Wybierz plik BBS",
                Filter     = "Excel BBS (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx"
            };
            if (fileDlg.ShowDialog() != true) return;

            try
            {
                var result   = BmmChecker.CheckAll(fileDlg.FileName);
                int failCount = new[] { result.R87, result.R95, result.R81, result.R83, result.R92 }
                    .Count(r => r.Status == "FAIL");

                doc.Editor.WriteMessage($"\nBMM: {failCount} błędów znaleziono — sprawdź okno wyników.\n");

                var resultDlg = new BmmResultsDialog(result, System.IO.Path.GetFileName(fileDlg.FileName));
                AcApp.ShowModalWindow(AcApp.MainWindow.Handle, resultDlg, false);
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\nBMM błąd: {ex.Message}\n");
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
            pso1.MessageForAdding = "\nZaznacz oknem wszystkie pręty H10 (zestawienie B1/B2): ";
            var sel1 = ed.GetSelection(pso1);
            if (sel1.Status != PromptStatus.OK) return;

            var barsB = ReadTemplateBarPositions(doc.Database, sel1.Value.GetObjectIds());
            if (barsB.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    "Nie wykryto prętów H10. Upewnij się że zaznaczono pręty ASD (rbcr_def_bar_bv).",
                    "ASD-GSETUP", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }
            SessionData.TemplateBarsB = barsB;

            // ── H12 bars (T1/T2) ─────────────────────────────────────────────────
            var pso2 = new PromptSelectionOptions();
            pso2.MessageForAdding = "\nZaznacz oknem wszystkie pręty H12 (zestawienie T1/T2): ";
            var sel2 = ed.GetSelection(pso2);
            if (sel2.Status != PromptStatus.OK) return;

            var barsT = ReadTemplateBarPositions(doc.Database, sel2.Value.GetObjectIds());
            if (barsT.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    "Nie wykryto prętów H12.",
                    "ASD-GSETUP", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }
            SessionData.TemplateBarsT = barsT;

            ed.WriteMessage($"\nGSETUP: H10={barsB.Count} prętów, H12={barsT.Count} prętów. Gotowy do ASD-GBOT/ASD-GTOP.\n");
            System.Windows.MessageBox.Show(
                $"Zarejestrowano szablony:\n  H10 (B1/B2): {barsB.Count} prętów [{string.Join(", ", System.Linq.Enumerable.Select(barsB.Keys, k => k + "mm"))}]\n  H12 (T1/T2): {barsT.Count} prętów",
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
                Title  = "Wybierz plik Punching",
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
                        "Nie znaleziono sekcji 'PLOT N' w arkuszu 'Punching Report to Calcs'.\n\n" +
                        "Sprawdź czy plik to raport punching w nowym formacie.",
                        "PXIE", System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                    return;
                }

                // Krok 2: wybór plotu (auto jeśli jeden, dialog jeśli wiele)
                PlotInfo selectedPlot;
                if (plots.Count == 1)
                {
                    selectedPlot = plots[0];
                    ed.WriteMessage($"\nPXIE: Auto-wybrano {selectedPlot}.\n");
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
                        "Brak wczytanych pali. Możliwe przyczyny:\n" +
                        "• plik nie został przeliczony w Excelu (otwórz i zapisz Ctrl+S)\n" +
                        "• sekcje w pliku są puste\n\n" +
                        "Sprawdź log w command line po szczegóły.",
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
                    $"\nPXIE: Wczytano {piles.Count} pali z {selectedPlot} " +
                    $"(INT:{selectedPlot.InternalCount} EDGE:{selectedPlot.EdgeCount} " +
                    $"CORNER:{selectedPlot.CornerCount}{reentrantPart}). Gotowy do Assign PH.\n");
                System.Windows.MessageBox.Show(
                    $"Wczytano {piles.Count} pali z {selectedPlot}.\nGotowy do Assign PH.",
                    "Wczytaj Punching", System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nPXIE błąd: {ex.Message}\n");
            }
        }

        [CommandMethod("ASD-PAA")]
        public void CmdAssignPH()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;

            if (SessionData.Piles == null || SessionData.Piles.Count == 0)
            {
                doc.Editor.WriteMessage("\nPAA: Najpierw użyj 'Wczytaj Punching' (ASD-PXIE).\n");
                System.Windows.MessageBox.Show("Najpierw wczytaj dane pali przyciskiem 'Wczytaj Punching'.",
                    "Assign PH", System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            try
            {
                PhAssigner.AssignAll(SessionData.Piles);
                SessionData.PhAssigned = true;

                doc.Editor.WriteMessage($"\nPAA: Przypisano PH dla {SessionData.Piles.Count} pali.\n");

                // Pokaż wyniki i zapytaj czy podpisać rysunek
                var dlg = new PhAssignResultsDialog(SessionData.Piles);
                AcApp.ShowModalWindow(AcApp.MainWindow.Handle, dlg, false);

                // Annotuj rysunek (znajdź kółka i podpisz PH)
                var res = DrawingAnnotator.Annotate(SessionData.Piles);
                doc.Editor.WriteMessage($"\nPAA: {res.Log.Replace("\n", " ")}");

                if (res.WrongDrawing)
                {
                    System.Windows.MessageBox.Show(
                        res.Log,
                        "Assign PH — zły rysunek",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                    return;
                }

                if (res.NotFound.Count > 0)
                {
                    System.Windows.MessageBox.Show(
                        $"Podpisano {res.Annotated.Count} pali na rysunku.\n\n" +
                        $"Nie znaleziono kółek/etykiet dla:\n{string.Join("\n", res.NotFound)}",
                        "Assign PH — wyniki", System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                }
                else if (res.Annotated.Count > 0)
                {
                    System.Windows.MessageBox.Show(
                        $"Podpisano {res.Annotated.Count} pali na rysunku.\n" +
                        $"Warstwy: {DrawingAnnotator.LayerPhText} (etykiety), {DrawingAnnotator.LayerPhHatch} (hatch).",
                        "Assign PH — OK", System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                }
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\nPAA błąd: {ex.Message}\n");
            }
        }

        [CommandMethod("ASD-PHR")]
        public void CmdPHReport()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;

            if (!SessionData.PhAssigned || SessionData.Piles == null)
            {
                doc.Editor.WriteMessage("\nPHR: Najpierw użyj Assign PH.\n");
                System.Windows.MessageBox.Show("Najpierw wykonaj Assign PH.",
                    "PH Report", System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            // Otwiera dialog z Assign PH (zawiera przycisk Eksportuj do Excel)
            var dlg = new PhAssignResultsDialog(SessionData.Piles);
            AcApp.ShowModalWindow(AcApp.MainWindow.Handle, dlg, false);
        }

        [CommandMethod("ASD-PHV")]
        public void CmdWalidujPH()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;

            if (!SessionData.PhAssigned || SessionData.Piles == null)
            {
                doc.Editor.WriteMessage("\nPHV: Najpierw wykonaj Assign PH.\n");
                return;
            }

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("PHV — WALIDACJA PH:");
            sb.AppendLine(new string('-', 40));

            // R77: brak EXCEED
            var exceed = SessionData.Piles.Where(p => p.PhAction == "EXCEED").ToList();
            if (exceed.Any())
                sb.AppendLine($"R77: FAIL — Util > 100%: {string.Join(", ", exceed.Select(p => p.PileId))}");
            else
                sb.AppendLine("R77: OK — Brak pali z Util > 100%");

            // R79: brak orphan (puste ApplicablePileIds)
            var orphan = SessionData.Piles.Where(p => p.ApplicablePileIds == null || p.ApplicablePileIds.Count == 0).ToList();
            if (orphan.Any())
                sb.AppendLine($"R79: FAIL — Orphan PH: {string.Join(", ", orphan.Select(p => p.PileId))}");
            else
                sb.AppendLine("R79: OK — Wszystkie pale mają ApplicablePileIds");

            // R27: duplikaty PileId
            var dupes = SessionData.Piles.GroupBy(p => p.PileId).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
            if (dupes.Any())
                sb.AppendLine($"R27: FAIL — Duplikaty: {string.Join(", ", dupes)}");
            else
                sb.AppendLine("R27: OK — Brak duplikatów Pile ID");

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
            if (string.IsNullOrWhiteSpace(newPrefix)) return rcTitle1;
            var m = Regex.Match(rcTitle1, @"REINFORCEMENT\s+DETAILS", RegexOptions.IgnoreCase);
            if (!m.Success) return rcTitle1;
            return newPrefix + " " + rcTitle1.Substring(m.Index);
        }

        private enum MsTextCategory { MainBottom, MainTop, Section, Ph, Detail }

        private static List<(MsTextCategory cat, double x, double y)> ScanModelSpaceTexts(Database db)
        {
            var result = new List<(MsTextCategory, double, double)>();
            var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;
            int totalTexts = 0, classified = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    string content = null;
                    string layer   = null;
                    double x = 0, y = 0;

                    var ent = tr.GetObject(id, OpenMode.ForRead);
                    if (ent is MText mt)
                    {
                        content = mt.Contents;
                        layer   = mt.Layer;
                        x = mt.Location.X;
                        y = mt.Location.Y;
                    }
                    else if (ent is DBText t)
                    {
                        content = t.TextString;
                        layer   = t.Layer;
                        x = t.Position.X;
                        y = t.Position.Y;
                    }
                    else continue;

                    if (string.IsNullOrEmpty(content)) continue;
                    totalTexts++;

                    string cat      = null;
                    MsTextCategory? msCat = null;

                    var mainMatch = MainLayerRx.Match(content);
                    if (mainMatch.Success)
                    {
                        string which = mainMatch.Groups[1].Value.ToUpperInvariant();
                        msCat = which == "BOTTOM" ? MsTextCategory.MainBottom : MsTextCategory.MainTop;
                        cat   = "MAIN_" + which;
                    }
                    else if (SectionRx.IsMatch(content)) { msCat = MsTextCategory.Section; cat = "SECTION"; }
                    else if (PhRx.IsMatch(content))      { msCat = MsTextCategory.Ph;      cat = "PH"; }
                    else if (DetailRx.IsMatch(content))  { msCat = MsTextCategory.Detail;  cat = "DETAIL"; }

                    if (msCat.HasValue)
                    {
                        result.Add((msCat.Value, x, y));
                        classified++;
                    }

                    bool interesting = cat != null
                        || content.IndexOf("DETAIL",  StringComparison.OrdinalIgnoreCase) >= 0
                        || content.IndexOf("SECTION", StringComparison.OrdinalIgnoreCase) >= 0
                        || content.IndexOf("LAYER",   StringComparison.OrdinalIgnoreCase) >= 0
                        || content.IndexOf("PH",      StringComparison.OrdinalIgnoreCase) >= 0;

                    if (interesting)
                    {
                        string snippet = content.Length > 80 ? content.Substring(0, 80) + "..." : content;
                        snippet = snippet.Replace("\n", "\\n").Replace("\r", "");
                        string catStr = cat ?? "(nie sklasyfikowano)";
                        ed.WriteMessage($"\nGAI-MS [{catStr}] layer='{layer}' pos=({x:F0},{y:F0}) | {snippet}");
                    }
                }
                tr.Commit();
            }
            ed.WriteMessage($"\nGAI-MS: scan zakonczony, totalTexts={totalTexts}, classified={classified}");
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

        private static string BuildDrawingSuffix(int? plotNumber, int layoutIdx0Based)
        {
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
                        if (!string.Equals(mt.Layer, layerName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        string contents = mt.Contents;
                        if (!contents.Contains("SLAB AREA")) continue;

                        var m1 = SlabAreaExtractRx.Match(contents);
                        var m2 = SlabPerimeterExtractRx.Match(contents);
                        var m3 = SlabThicknessExtractRx.Match(contents);

                        if (m1.Success) result[KeySlabArea]      = m1.Groups[1].Value;
                        if (m2.Success) result[KeySlabPerimeter] = m2.Groups[1].Value;
                        if (m3.Success) result[KeySlabThickness] = m3.Groups[1].Value;

                        var vMatch = ConcreteVolumeExtractRx.Match(contents);
                        if (vMatch.Success)
                        {
                            result[KeyConcreteVolume]     = vMatch.Groups[2].Value;
                            result[KeyConcreteVolumeTail] = vMatch.Groups[4].Value;
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
                        if (!string.Equals(mt.Layer, layerName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        string contents = mt.Contents;
                        if (!contents.Contains("SLAB AREA")) continue;

                        string newContents = contents;

                        if (values.TryGetValue(KeySlabArea, out var vArea) && !string.IsNullOrEmpty(vArea))
                            newContents = SlabAreaReplaceRx.Replace(newContents, "${1}" + vArea + "${3}");
                        if (values.TryGetValue(KeySlabPerimeter, out var vPer) && !string.IsNullOrEmpty(vPer))
                            newContents = SlabPerimeterReplaceRx.Replace(newContents, "${1}" + vPer + "${3}");
                        if (values.TryGetValue(KeySlabThickness, out var vTh) && !string.IsNullOrEmpty(vTh))
                            newContents = SlabThicknessReplaceRx.Replace(newContents, "${1}" + vTh + "${3}");

                        // CONCRETE VOLUME — nadpisuje wartość liczbową i tail
                        if (values.TryGetValue(KeyConcreteVolume, out var vVol) && !string.IsNullOrEmpty(vVol))
                        {
                            string newTail = values.TryGetValue(KeyConcreteVolumeTail, out var t) ? t : "";
                            string newTailEscaped = newTail.Replace("$", "$$");
                            newContents = ConcreteVolumeExtractRx.Replace(newContents,
                                              "${1}" + vVol + "${3}" + newTailEscaped);
                        }

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
            Dictionary<string, string> autoTitle3Map,
            Dictionary<string, string> autoScalesMap,
            string nowDate)
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

                            // Specjalna obsługa TITLE_1: podmień prefix przed "REINFORCEMENT DETAILS"
                            if (!string.IsNullOrEmpty(gaTitlePrefix) &&
                                string.Equals(att.Tag, "TITLE_1", StringComparison.OrdinalIgnoreCase))
                            {
                                newVal = ReplaceRcTitlePrefix(att.TextString, gaTitlePrefix);
                            }
                            // DRAWING_NUMBER: buduj suffix wg numeru plotu i indeksu layoutu
                            else if (!string.IsNullOrEmpty(gaDrawingPrefix) &&
                                     string.Equals(att.Tag, "DRAWING_NUMBER", StringComparison.OrdinalIgnoreCase))
                            {
                                string suffix = BuildDrawingSuffix(currentLayoutPlot, currentLayoutIdx);
                                newVal = gaDrawingPrefix + "-" + suffix;
                            }
                            // TITLE_3: auto-wykrycie z viewportów
                            else if (autoTitle3Map != null &&
                                     string.Equals(att.Tag, "TITLE_3", StringComparison.OrdinalIgnoreCase))
                            {
                                if (autoTitle3Map.TryGetValue(layout.LayoutName, out string t3)
                                    && !string.IsNullOrEmpty(t3))
                                {
                                    newVal = t3;
                                }
                            }
                            // SCALE: auto z viewportów
                            else if (autoScalesMap != null &&
                                     string.Equals(att.Tag, "SCALE", StringComparison.OrdinalIgnoreCase))
                            {
                                if (autoScalesMap.TryGetValue(layout.LayoutName, out string sc)
                                    && !string.IsNullOrEmpty(sc))
                                {
                                    newVal = sc;
                                }
                            }
                            // DATE: aktualny miesiąc + rok
                            else if (!string.IsNullOrEmpty(nowDate) &&
                                     string.Equals(att.Tag, "DATE", StringComparison.OrdinalIgnoreCase))
                            {
                                newVal = nowDate;
                            }
                            // Standardowa obsługa CLIENT_* / PROJ_*
                            else if (tagsSet.Contains(att.Tag))
                            {
                                src.TryGetValue(att.Tag, out newVal);
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
