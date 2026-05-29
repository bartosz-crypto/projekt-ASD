using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;

[assembly: ExtensionApplication(typeof(AsdRcSlab.App))]

namespace AsdRcSlab
{
    public class App : IExtensionApplication
    {
        public void Initialize()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.Editor.WriteMessage("\nASD RC SLAB v4.4 loaded. Wpisz ASD-PROJ aby zaczac.\n");
            }

            try
            {
                // Spróbuj zbudować ribbon od razu — jeśli API już gotowe.
                if (ComponentManager.Ribbon != null)
                {
                    RibbonBuilder.Build();
                    return;
                }

                // Ribbon API jeszcze nie gotowe (typowe podczas auto-load w ASD 2015).
                // Podepnij handler na event ItemInitialized — wywoła się gdy ribbon
                // tab/element zostanie zainicjalizowany.
                ComponentManager.ItemInitialized += OnComponentManagerItemInitialized;

                // Fallback — jeśli ItemInitialized nie odpali, Application.Idle
                // sprawdza przy każdym cyklu pętli komunikatów czy ribbon jest ready.
                Autodesk.AutoCAD.ApplicationServices.Application.Idle += OnApplicationIdle;
            }
            catch (System.Exception ex)
            {
                doc?.Editor.WriteMessage($"\nBlad ladowania ribbon: {ex.Message}\n");
            }
        }

        private void OnComponentManagerItemInitialized(object sender, RibbonItemEventArgs e)
        {
            if (ComponentManager.Ribbon == null) return;

            ComponentManager.ItemInitialized -= OnComponentManagerItemInitialized;
            Autodesk.AutoCAD.ApplicationServices.Application.Idle -= OnApplicationIdle;

            try
            {
                RibbonBuilder.Build();
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
                    .MdiActiveDocument?.Editor
                    .WriteMessage($"\nAsdRcSlab ribbon error: {ex.Message}\n");
            }
        }

        private void OnApplicationIdle(object sender, System.EventArgs e)
        {
            if (ComponentManager.Ribbon == null) return;

            Autodesk.AutoCAD.ApplicationServices.Application.Idle -= OnApplicationIdle;
            ComponentManager.ItemInitialized -= OnComponentManagerItemInitialized;

            try
            {
                RibbonBuilder.Build();
            }
            catch (System.Exception ex)
            {
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
