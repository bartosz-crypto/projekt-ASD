using Autodesk.AutoCAD.Runtime;
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
            var dlg = new ReinforcementStubDialog(dia: 10, layer: "B1/B2", asBase: 393);
            AcApp.ShowModalWindow(AcApp.MainWindow.Handle, dlg, false);
        }

        [CommandMethod("ASD-GTOP")]
        public void CmdGenerujT1T2()
        {
            var dlg = new ReinforcementStubDialog(dia: 12, layer: "T1/T2", asBase: 565);
            AcApp.ShowModalWindow(AcApp.MainWindow.Handle, dlg, false);
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
