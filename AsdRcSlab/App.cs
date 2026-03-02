using Autodesk.AutoCAD.Runtime;
using System;

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
                RibbonBuilder.Build();
            }
            catch (System.Exception ex)
            {
                doc?.Editor.WriteMessage($"\nBlad ladowania ribbon: {ex.Message}\n");
            }
        }

        public void Terminate() { }
    }
}
