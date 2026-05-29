using System.IO;
using Autodesk.AutoCAD.Customization;

namespace AsdRcSlab
{
    /// <summary>
    /// Generuje partial CUIx programowo przez CustomizationSection API
    /// i powiazuje ribbon tab z workspace - zeby ASD 2015 renderowal
    /// ribbon (ASD renderuje z workspace, nie z
    /// ComponentManager.Ribbon.Tabs bezposrednio).
    ///
    /// Sygnatury API zweryfikowane przez reflection na AcCui.dll v20.0
    /// (AutoCAD 2015) - p127.
    /// </summary>
    public static class CuixBuilder
    {
        // Definicja przyciskow: (panelTitle, buttonText, command)
        private static readonly (string panel, string text, string cmd)[] Buttons = new[]
        {
            ("TITLE BLOCK",        "Copy from GA",    "ASD-GAI"),
            ("TITLE BLOCK",        "Sheet Numbering", "ASD-RCN"),
            ("PH CONDITIONS",      "Load Punching",   "ASD-PXIE"),
            ("PH CONDITIONS",      "Assign PH",       "ASD-PAA"),
            ("PH CONDITIONS",      "PH Report",       "ASD-PHR"),
            ("PH CONDITIONS",      "Waliduj PH",      "ASD-PHV"),
            ("REINFORCEMENT MAPS", "Import Maps",     "ASD-IMR"),
            ("BBS",                "Bar Calculator",  "ASD-BBC"),
            ("BBS",                "BBS Write",       "ASD-BBS-WRITE"),
        };

        private const string TabId = "ASD_RC_SLAB_TAB";
        private const string TabName = "ASD RC SLAB";
        private const string MenuGroupName = "ASDRCSLAB";
        private const string MacroGroupName = "ASDRCSLAB_MACROS";

        /// <summary>
        /// Sciezka do partial CUIx w folderze bundle Contents (obok DLL).
        /// </summary>
        public static string GetCuixPath()
        {
            var dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var dir = Path.GetDirectoryName(dllPath);
            return Path.Combine(dir, "AsdRcSlab.cuix");
        }

        public static void BuildAndLoad()
        {
            App.DiagLog("[CUIX] BuildAndLoad START");
            string cuixPath = GetCuixPath();
            App.DiagLog($"[CUIX] target path: {cuixPath}");

            try
            {
                // 1. Stworz lub otworz partial CUIx
                CustomizationSection cs;
                if (File.Exists(cuixPath))
                {
                    App.DiagLog("[CUIX] opening existing CUIx");
                    cs = new CustomizationSection(cuixPath);
                }
                else
                {
                    App.DiagLog("[CUIX] creating new CUIx");
                    cs = new CustomizationSection();
                    cs.MenuGroupName = MenuGroupName;
                    cs.EnsureMenuHasDefaults();
                    cs.SaveAs(cuixPath);
                    // reopen po save zeby miec spojny stan na dysku
                    cs = new CustomizationSection(cuixPath);
                }

                App.DiagLog($"[CUIX] MenuGroupName = {cs.MenuGroupName}");

                // 2. RibbonRoot
                var ribbonRoot = cs.MenuGroup.RibbonRoot;
                App.DiagLog($"[CUIX] RibbonRoot tabs before = {ribbonRoot.RibbonTabSources.Count}");

                // Usun istniejacy tab/panele o naszych Id (idempotent rebuild)
                for (int i = ribbonRoot.RibbonTabSources.Count - 1; i >= 0; i--)
                {
                    if (ribbonRoot.RibbonTabSources[i].ElementID == TabId)
                    {
                        ribbonRoot.RibbonTabSources.Remove(i);
                        App.DiagLog("[CUIX] removed existing tab source");
                    }
                }
                for (int i = ribbonRoot.RibbonPanelSources.Count - 1; i >= 0; i--)
                {
                    var pid = ribbonRoot.RibbonPanelSources[i].ElementID ?? "";
                    if (pid.StartsWith("ASD_PANEL_"))
                        ribbonRoot.RibbonPanelSources.Remove(i);
                }

                // 3. RibbonTabSource
                var tabSource = new RibbonTabSource(ribbonRoot);
                tabSource.Name = TabName;
                tabSource.Text = TabName;
                tabSource.ElementID = TabId;
                ribbonRoot.RibbonTabSources.Add(tabSource);
                App.DiagLog("[CUIX] tab source created");

                // 4. MacroGroup (znajdz lub stworz)
                var macroGroup = cs.MenuGroup.FindMacroGroup(MacroGroupName);
                if (macroGroup == null)
                {
                    macroGroup = new MacroGroup(MacroGroupName, cs.MenuGroup);
                    if (cs.MenuGroup.FindMacroGroup(MacroGroupName) == null)
                        cs.MenuGroup.MacroGroups.Add(macroGroup);
                    App.DiagLog("[CUIX] macro group created");
                }

                // 5. Panele + przyciski
                string lastPanel = null;
                RibbonPanelSource currentPanelSource = null;
                RibbonRow currentRow = null;

                foreach (var (panel, text, cmd) in Buttons)
                {
                    if (panel != lastPanel)
                    {
                        currentPanelSource = new RibbonPanelSource(ribbonRoot);
                        currentPanelSource.Name = panel;
                        currentPanelSource.Text = panel;
                        currentPanelSource.ElementID = $"ASD_PANEL_{panel.Replace(" ", "_")}";
                        ribbonRoot.RibbonPanelSources.Add(currentPanelSource);

                        // Panel reference w tab
                        var panelRef = new RibbonPanelSourceReference(tabSource);
                        panelRef.PanelId = currentPanelSource.ElementID;
                        tabSource.Items.Add(panelRef);

                        // Row w panelu (ctor bezparametrowy - parent ustawiany przez Items.Add)
                        currentRow = new RibbonRow();
                        currentPanelSource.Items.Add(currentRow);

                        lastPanel = panel;
                        App.DiagLog($"[CUIX] panel source created: {panel}");
                    }

                    // MenuMacro: ctor (parent, name, command, tag). Brak settera Name/Command.
                    string macroId = $"ASD_MACRO_{cmd.Replace("-", "_")}";
                    var existingMacro = cs.FindMenuMacro(macroId);
                    MenuMacro macro = existingMacro ?? new MenuMacro(
                        macroGroup, text, $"^C^C_{cmd}\n", macroId);
                    macro.ElementID = macroId;

                    // Przycisk (parent = row)
                    var btn = new RibbonCommandButton(currentRow);
                    btn.MacroID = macro.ElementID;
                    btn.Text = text;
                    btn.ButtonStyle = RibbonButtonStyle.LargeWithText;
                    currentRow.Items.Add(btn);

                    App.DiagLog($"[CUIX]   button: {text} -> {cmd}");
                }

                // 6. Powiazanie tab z workspace
                // === p127b: DIAG Workspace API ===
                App.DiagLog($"[CUIX] workspaces count = {cs.Workspaces.Count}");

                // Pobierz nazwe aktywnego workspace z systemu
                string activeWs = null;
                try
                {
                    activeWs = Autodesk.AutoCAD.ApplicationServices.Application
                        .GetSystemVariable("WSCURRENT") as string;
                }
                catch { }
                App.DiagLog($"[CUIX] active workspace (WSCURRENT) = {activeWs}");

                // Reflection: konstruktor Workspace + property WorkspaceRibbonRoot / WSRibbonTabSourceReference
                try
                {
                    var wsType = typeof(Workspace);
                    App.DiagLog("[CUIX-API] Workspace ctors:");
                    foreach (var c in wsType.GetConstructors())
                        App.DiagLog($"[CUIX-API]   ctor({string.Join(", ", System.Linq.Enumerable.Select(c.GetParameters(), p => p.ParameterType.Name + " " + p.Name))})");
                    App.DiagLog("[CUIX-API] Workspace properties:");
                    foreach (var p in wsType.GetProperties())
                        App.DiagLog($"[CUIX-API]   {p.PropertyType.Name} {p.Name}");

                    var wsTabRefType = typeof(WSRibbonTabSourceReference);
                    App.DiagLog("[CUIX-API] WSRibbonTabSourceReference ctors:");
                    foreach (var c in wsTabRefType.GetConstructors())
                        App.DiagLog($"[CUIX-API]   ctor({string.Join(", ", System.Linq.Enumerable.Select(c.GetParameters(), p => p.ParameterType.Name + " " + p.Name))})");
                    App.DiagLog("[CUIX-API] WSRibbonTabSourceReference properties:");
                    foreach (var p in wsTabRefType.GetProperties())
                        App.DiagLog($"[CUIX-API]   {p.PropertyType.Name} {p.Name}");

                    var wsRootType = typeof(WSRibbonRoot);
                    App.DiagLog("[CUIX-API] WSRibbonRoot properties:");
                    foreach (var p in wsRootType.GetProperties())
                        App.DiagLog($"[CUIX-API]   {p.PropertyType.Name} {p.Name}");
                }
                catch (System.Exception ex)
                {
                    App.DiagLog($"[CUIX-API] reflection EXCEPTION: {ex.Message}");
                }

                // === p127b: stworz workspace z tab reference ===
                // Sygnatury zweryfikowane reflection (AcCui 20.0): ctor Workspace
                // to (CustomizationSection parent, String name); property to
                // WorkspaceRibbonRoot (nie WorkspaceRoot.RibbonRoot). Tab podpinamy
                // helperem MergeOrAddTabToWorkspace - buduje WSRibbonTabSourceReference
                // poprawnie (ElementID + MenuGroup) bez recznego zgadywania.
                try
                {
                    string wsName = !string.IsNullOrEmpty(activeWs) ? activeWs : "A2";

                    // Idempotent: znajdz lub stworz workspace o tej nazwie
                    Workspace targetWs = null;
                    foreach (Workspace w in cs.Workspaces)
                    {
                        if (w.Name == wsName) { targetWs = w; break; }
                    }

                    if (targetWs == null)
                    {
                        targetWs = new Workspace(cs, wsName);
                        if (cs.Workspaces.IndexOfWorkspaceName(wsName) < 0)
                            cs.Workspaces.Add(targetWs);
                        App.DiagLog($"[CUIX] workspace '{wsName}' created in partial");
                    }
                    else
                    {
                        App.DiagLog($"[CUIX] workspace '{wsName}' already in partial");
                    }

                    // Podepnij tab do workspace (merge przy CUILOAD do glownego A2 ASD)
                    var refs = cs.MergeOrAddTabToWorkspace(targetWs, tabSource);
                    int refCount = refs != null ? refs.Count : -1;
                    App.DiagLog($"[CUIX] tab merged to workspace '{wsName}', refs = {refCount}");
                }
                catch (System.Exception ex)
                {
                    App.DiagLog($"[CUIX] workspace creation EXCEPTION {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
                }

                // 7. Zapisz CUIx
                bool saved = cs.Save();
                App.DiagLog($"[CUIX] saved = {saved}");

                // 8. Zaladuj partial menu do biezacej sesji
                var doc = Autodesk.AutoCAD.ApplicationServices.Application
                    .DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application
                        .SetSystemVariable("FILEDIA", 0);
                    doc.SendStringToExecute(
                        $"_.CUILOAD \"{cuixPath}\"\n", true, false, false);
                    App.DiagLog("[CUIX] CUILOAD sent");
                    Autodesk.AutoCAD.ApplicationServices.Application
                        .SetSystemVariable("FILEDIA", 1);
                }
                else
                {
                    App.DiagLog("[CUIX] brak active document - CUILOAD pominiety");
                }

                App.DiagLog("[CUIX] BuildAndLoad DONE");
            }
            catch (System.Exception ex)
            {
                App.DiagLog($"[CUIX] EXCEPTION {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            }
        }
    }
}
