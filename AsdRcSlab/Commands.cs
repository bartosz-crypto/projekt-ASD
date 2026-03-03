using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Newtonsoft.Json;
using System.IO;
using System.Linq;

namespace AsdRcSlab
{
    public class Commands
    {
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
        public void CmdWczytajGA()
        {
            System.Windows.MessageBox.Show(
                "GAI — Wczytaj GA będzie dostępne w Sprint 2.\n\nNa razie wprowadź dane ręcznie przez 'Nowy Projekt'.",
                "ASD RC SLAB — Wczytaj GA",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
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

            var fileDlg = new Microsoft.Win32.OpenFileDialog
            {
                Title  = "Wybierz plik Punching",
                Filter = "Excel (*.xlsx)|*.xlsx"
            };
            if (fileDlg.ShowDialog() != true) return;

            try
            {
                // Krok 1: pobierz listę arkuszy i poproś użytkownika o wybór
                var sheets = PunchingParser.GetSheetNames(fileDlg.FileName);
                if (sheets.Count == 0)
                {
                    doc.Editor.WriteMessage("\nPXIE: Plik nie zawiera żadnych arkuszy.\n");
                    return;
                }

                var sheetDlg = new SheetPickerDialog(sheets);
                if (AcApp.ShowModalWindow(AcApp.MainWindow.Handle, sheetDlg, false) != true) return;

                string selectedSheet = sheetDlg.SelectedSheet;

                // Krok 2: parsuj wybrany arkusz
                string log;
                var piles = PunchingParser.Parse(fileDlg.FileName, selectedSheet, out log);

                if (piles.Count == 0)
                {
                    doc.Editor.WriteMessage($"\nPXIE: Nie wczytano pali.\n{log}\n");
                    System.Windows.MessageBox.Show(
                        $"Nie znaleziono danych pali w arkuszu '{selectedSheet}'.\n\nLog parsera:\n{log}",
                        "Wczytaj Punching", System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                    return;
                }

                SessionData.Piles = piles;
                doc.Editor.WriteMessage($"\nPXIE: Wczytano {piles.Count} pali z arkusza '{selectedSheet}'. Gotowy do Assign PH.\n");
                System.Windows.MessageBox.Show(
                    $"Wczytano {piles.Count} pali z arkusza '{selectedSheet}'.\nGotowy do Assign PH.",
                    "Wczytaj Punching", System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\nPXIE błąd: {ex.Message}\n");
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
