using Autodesk.AutoCAD.Ribbon;
using Autodesk.Windows;
using System.Windows.Controls;

namespace AsdRcSlab
{
    public static class RibbonBuilder
    {
        public static void Build()
        {
            App.DiagLog("[Build] START");
            RibbonControl ribbon = ComponentManager.Ribbon;
            if (ribbon == null)
            {
                App.DiagLog("[Build] ribbon NULL - return");
                return;
            }

            App.DiagLog($"[Build] ribbon ready, existing tabs count = {ribbon.Tabs.Count}");

            // Usun istniejaca zakladke jesli juz istnieje
            bool existsBefore = false;
            foreach (RibbonTab existing in ribbon.Tabs)
            {
                if (existing.Id == "ASD_RC_SLAB_TAB")
                {
                    existsBefore = true;
                    ribbon.Tabs.Remove(existing);
                    break;
                }
            }
            App.DiagLog($"[Build] tab with same Id already present = {existsBefore} (removed if true)");

            RibbonTab tab = new RibbonTab
            {
                Title = "ASD RC SLAB",
                Id = "ASD_RC_SLAB_TAB"
            };

            tab.Panels.Add(CreatePanel("TITLE BLOCK",
                new[] {
                    ("Copy from GA",      "ASD-GAI", "Import attributes from GA to RC title blocks"),
                    ("Sheet Numbering",   "ASD-RCN", "Auto-fill TITLE_3, SCALE, DATE + rename layouts")
                }, columnsPerRow: 1));

            tab.Panels.Add(CreatePanel("PH CONDITIONS",
                new[] {
                    ("Load Punching", "ASD-PXIE", "Import PUNCHING_NEW_TEMPLATE_v2.xlsx"),
                    ("Assign PH",     "ASD-PAA",  "Assigns PH1-PH9 and generates detail titles"),
                    ("PH Report",        "ASD-PHR",  "Generuje PH_Report.xlsx"),
                    ("Waliduj PH",       "ASD-PHV",  "Sprawdza R77, R79, duplikaty")
                }, columnsPerRow: 2));

            tab.Panels.Add(CreatePanel("REINFORCEMENT MAPS",
                new[] {
                    ("Import Maps", "ASD-IMR", "Import reinforcement maps (TOP T1/T2, BOT B1/B2) from external DXF/DWG")
                }, columnsPerRow: 1));

            tab.Panels.Add(CreatePanel("BBS",
                new[] {
                    ("Bar Calculator", "ASD-BBC",
                     "Calculate bar cutting lengths from xlsx export (BS 8666:2020). "
                     + "Input: bar export xlsx; output: <name>_calculated.xlsx "
                     + "with raw + final lengths."),
                    ("BBS Write", "ASD-BBS-WRITE",
                     "Write calculated bar lengths to target BBS file (.xls/.xlsx). "
                     + "Clears existing data in BOTTOM/TOP LAYER sections, writes "
                     + "new rows with proper formatting (code 00 → STR, link bars "
                     + "skip C/D columns). Creates .bak backup before write.")
                }, columnsPerRow: 1));

            App.DiagLog($"[Build] about to add tab, tab.Id = {tab.Id}, tab.Title = {tab.Title}, panels = {tab.Panels.Count}");
            ribbon.Tabs.Add(tab);
            tab.IsActive = true;
            App.DiagLog($"[Build] tab added, new tabs count = {ribbon.Tabs.Count}, IsActive = {tab.IsActive}, DONE");

            // === p124: wymuszenie refresh ribbon UI ===
            // ASD 2015 nie renderuje programowo dodanego tab bez wymuszenia odswiezenia.
            try
            {
                // 1. UpdateLayout na ribbon control (WPF-level refresh)
                var ribbonControl = ComponentManager.Ribbon;
                if (ribbonControl != null)
                {
                    ribbonControl.UpdateLayout();
                    App.DiagLog("[Build] ribbon.UpdateLayout() called");
                }

                // 2. Wymuszenie re-aktywacji tab zeby ribbon przerysowal
                //    (czasem ustawienie IsActive na inny tab i z powrotem wymusza redraw)
                tab.IsVisible = true;
                tab.IsActive = true;
                App.DiagLog("[Build] tab IsVisible+IsActive re-set");
            }
            catch (System.Exception ex)
            {
                App.DiagLog($"[Build] refresh EXCEPTION: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private static RibbonPanel CreatePanel(
            string title,
            (string label, string cmd, string tooltip)[] buttons,
            int columnsPerRow = 1)
        {
            RibbonPanelSource source = new RibbonPanelSource { Title = title };
            RibbonRowPanel row = new RibbonRowPanel();

            for (int i = 0; i < buttons.Length; i++)
            {
                var (label, cmd, tooltip) = buttons[i];
                RibbonButton btn = new RibbonButton
                {
                    Text = label,
                    CommandHandler = new RibbonCommandHandler(cmd),
                    CommandParameter = cmd,
                    ShowText = true,
                    ShowImage = false,
                    Size = RibbonItemSize.Standard,
                    Width = 150,
                    MinWidth = 150,
                    ToolTip = tooltip
                };
                row.Items.Add(btn);

                bool isLast = (i == buttons.Length - 1);
                bool endOfRow = ((i + 1) % columnsPerRow == 0);
                if (endOfRow && !isLast)
                    row.Items.Add(new RibbonRowBreak());
            }

            source.Items.Add(row);

            RibbonPanel panel = new RibbonPanel { Source = source };
            return panel;
        }
    }
}
