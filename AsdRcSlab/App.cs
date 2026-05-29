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

        public void Initialize()
        {
            DiagLog("=== Initialize() START ===");

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.Editor.WriteMessage("\nASD RC SLAB v4.4 loaded. Wpisz ASD-PROJ aby zaczac.\n");
            }

            try
            {
                // Spróbuj zbudować ribbon od razu — jeśli API już gotowe.
                var ribbon = ComponentManager.Ribbon;
                DiagLog($"Initialize: ComponentManager.Ribbon = {(ribbon == null ? "NULL" : "READY")}");

                if (ribbon != null)
                {
                    DiagLog("Initialize: building ribbon synchronously");
                    RibbonBuilder.Build();          // stara metoda (zostaje do porownania)
                    CuixBuilder.BuildAndLoad();     // nowa metoda - CUIx + workspace
                    DiagLog("Initialize: synchronous Build() DONE");
                    return;
                }

                // Ribbon API jeszcze nie gotowe (typowe podczas auto-load w ASD 2015).
                // Podepnij handler na event ItemInitialized — wywoła się gdy ribbon
                // tab/element zostanie zainicjalizowany.
                DiagLog("Initialize: ribbon null - subscribing ItemInitialized + Idle");
                ComponentManager.ItemInitialized += OnComponentManagerItemInitialized;

                // Fallback — jeśli ItemInitialized nie odpali, Application.Idle
                // sprawdza przy każdym cyklu pętli komunikatów czy ribbon jest ready.
                Autodesk.AutoCAD.ApplicationServices.Application.Idle += OnApplicationIdle;
            }
            catch (System.Exception ex)
            {
                DiagLog($"Initialize: EXCEPTION {ex.GetType().Name}: {ex.Message}");
                doc?.Editor.WriteMessage($"\nBlad ladowania ribbon: {ex.Message}\n");
            }
        }

        private void OnComponentManagerItemInitialized(object sender, RibbonItemEventArgs e)
        {
            var ready = ComponentManager.Ribbon != null;
            DiagLog($"ItemInitialized fired: ribbon = {(ready ? "READY" : "NULL")}");
            if (!ready) return;

            ComponentManager.ItemInitialized -= OnComponentManagerItemInitialized;
            Autodesk.AutoCAD.ApplicationServices.Application.Idle -= OnApplicationIdle;

            try
            {
                DiagLog("ItemInitialized: calling Build()");
                RibbonBuilder.Build();          // stara metoda (zostaje do porownania)
                CuixBuilder.BuildAndLoad();     // nowa metoda - CUIx + workspace
                DiagLog("ItemInitialized: Build() DONE");
            }
            catch (System.Exception ex)
            {
                DiagLog($"ItemInitialized: EXCEPTION {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
                    .MdiActiveDocument?.Editor
                    .WriteMessage($"\nAsdRcSlab ribbon error: {ex.Message}\n");
            }
        }

        private void OnApplicationIdle(object sender, System.EventArgs e)
        {
            var ready = ComponentManager.Ribbon != null;
            // Loguj tylko gdy ribbon ready — Idle odpala setki razy, nie zalewamy pliku.
            if (!ready) return;

            DiagLog("Idle fired: ribbon READY");
            Autodesk.AutoCAD.ApplicationServices.Application.Idle -= OnApplicationIdle;
            ComponentManager.ItemInitialized -= OnComponentManagerItemInitialized;

            try
            {
                DiagLog("Idle: calling Build()");
                RibbonBuilder.Build();          // stara metoda (zostaje do porownania)
                CuixBuilder.BuildAndLoad();     // nowa metoda - CUIx + workspace
                DiagLog("Idle: Build() DONE");
            }
            catch (System.Exception ex)
            {
                DiagLog($"Idle: EXCEPTION {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
                    .MdiActiveDocument?.Editor
                    .WriteMessage($"\nAsdRcSlab ribbon error: {ex.Message}\n");
            }
        }

        public void Terminate()
        {
            try
            {
                ComponentManager.ItemInitialized -= OnComponentManagerItemInitialized;
                Autodesk.AutoCAD.ApplicationServices.Application.Idle -= OnApplicationIdle;
            }
            catch { /* ignore */ }
        }
    }
}
