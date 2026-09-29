using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;

[assembly: ExtensionApplication(typeof(AsdRcSlab.App))]

namespace AsdRcSlab
{
    public class App : IExtensionApplication
    {
        // p123 DIAG — logging ribbon build timing do pliku (editor moze nie istniec przy starcie)
        private static readonly string DiagLogPath =
            System.IO.Path.Combine(
                System.Environment.GetEnvironmentVariable("TEMP") ?? @"C:\Temp",
                "AsdRcSlab-ribbon-diag.log");

        internal static void DiagLog(string msg)
        {
            try
            {
                System.IO.File.AppendAllText(DiagLogPath,
                    $"{System.DateTime.Now:HH:mm:ss.fff}  {msg}{System.Environment.NewLine}");
            }
            catch { /* ignore */ }
        }

        // Wykrywa, czy biezacy proces to AutoCAD Structural Detailing (ASD),
        // a nie zwykly AutoCAD. ASD startuje tym samym acad.exe, ale ze znacznikami
        // w linii polecen: /product ASD, /p "AutoCAD Structural Detailing", ASD.arx.
        // Dziala tak samo na 2013/2014/2015.
        internal static bool IsAsd()
        {
            try
            {
                string cmd = (System.Environment.CommandLine ?? "").ToUpperInvariant();
                return cmd.Contains("/PRODUCT ASD")
                    || cmd.Contains("ASD.ARX")
                    || cmd.Contains("STRUCTURAL DETAILING");
            }
            catch { return false; }
        }

        private bool _ribbonScheduled = false;

        public void Initialize()
        {
            DiagLog("=== Initialize() START ===");

            // Wtyczka ma dzialac TYLKO w ASD, nie w zwyklym AutoCAD.
            if (!IsAsd())
            {
                DiagLog("Initialize: to NIE jest ASD (/product != ASD) - wtyczka nieaktywna w zwyklym AutoCAD");
                return;
            }
            DiagLog("Initialize: wykryto ASD (AutoCAD Structural Detailing)");

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            doc?.Editor.WriteMessage(
                $"\nASD RC SLAB v{RibbonBuilder.Version} (build {RibbonBuilder.BuildStamp()}) loaded. Wpisz ASD-PROJ aby zaczac.\n");

            // Jesli wstazka jest wylaczona (klasyczny workspace ASD) - wlacz ja.
            try
            {
                if (ComponentManager.Ribbon == null)
                {
                    doc?.SendStringToExecute("_.RIBBON ", true, false, false);
                    DiagLog("Initialize: wstazka byla wylaczona - wyslano _.RIBBON");
                }
            }
            catch (System.Exception exr) { DiagLog($"Initialize: _.RIBBON err: {exr.Message}"); }

            // Wstazke budujemy ZAWSZE z opoznieniem, przez kolejke polecen
            // (ASD-RIBBONBUILD) - dokladnie tak jak dziala reczne ASD-RIBBON,
            // ktore okazalo sie niezawodne. Budowanie synchroniczne przy starcie
            // bywalo za wczesnie (workspace ASD jeszcze niegotowy) i zakladka sie
            // nie pojawiala. Czekamy na pierwszy Idle z gotowa wstazka i dokumentem.
            Autodesk.AutoCAD.ApplicationServices.Application.Idle += OnIdleBuildRibbon;
            DiagLog("Initialize: zaplanowano odlozone zbudowanie wstazki (Idle)");
        }

        private void OnIdleBuildRibbon(object sender, System.EventArgs e)
        {
            if (_ribbonScheduled) return;
            if (ComponentManager.Ribbon == null) return;   // czekaj az wstazka wstanie
            var doc = Autodesk.AutoCAD.ApplicationServices.Application
                .DocumentManager.MdiActiveDocument;
            if (doc == null) return;                        // czekaj na aktywny dokument

            _ribbonScheduled = true;
            Autodesk.AutoCAD.ApplicationServices.Application.Idle -= OnIdleBuildRibbon;

            try
            {
                // To samo, co reczne ASD-RIBBON: zbuduj zakladke w kontekscie
                // polecenia, gdy ASD jest juz w pelni gotowe.
                DiagLog("Idle: wstazka+dokument gotowe -> wysylam ASD-RIBBONBUILD (odlozone)");
                doc.SendStringToExecute("ASD-RIBBONBUILD ", true, false, false);
            }
            catch (System.Exception ex)
            {
                DiagLog($"Idle: ASD-RIBBONBUILD send EXCEPTION: {ex.Message}");
            }
        }

        public void Terminate()
        {
            try { Autodesk.AutoCAD.ApplicationServices.Application.Idle -= OnIdleBuildRibbon; }
            catch { /* ignore */ }
        }
    }
}
