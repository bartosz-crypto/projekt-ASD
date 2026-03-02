using Autodesk.AutoCAD.Runtime;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Newtonsoft.Json;
using System.IO;

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
            doc.Editor.WriteMessage("\nTODO: Generuj B1/B2 — siatka dolna H10@200\n");
        }

        [CommandMethod("ASD-GTOP")]
        public void CmdGenerujT1T2()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Generuj T1/T2 — siatka gorna H12@200\n");
        }

        [CommandMethod("ASD-BMM")]
        public void CmdOznaczPrety()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Oznacz Prety — walidacja BBS R87/R95/R81/R83\n");
        }

        [CommandMethod("ASD-LAP")]
        public void CmdZakladyAuto()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Zaklady Auto — kalkulator Note 16\n");
        }

        // ── PANEL 3: PH CONDITIONS ────────────────────────────────────────────

        [CommandMethod("ASD-PXIE")]
        public void CmdWczytajPunching()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Wczytaj Punching — import PUNCHING_NEW_TEMPLATE_v2.xlsx\n");
        }

        [CommandMethod("ASD-PAA")]
        public void CmdAssignPH()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Assign PH — logika PH1-PH9 + REENTRANT\n");
        }

        [CommandMethod("ASD-PHR")]
        public void CmdPHReport()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: PH Report — eksport PH_Report.xlsx\n");
        }

        [CommandMethod("ASD-PHV")]
        public void CmdWalidujPH()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nTODO: Waliduj PH — R77, R79, duplikaty\n");
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
