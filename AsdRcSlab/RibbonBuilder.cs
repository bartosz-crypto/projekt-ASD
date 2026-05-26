using Autodesk.AutoCAD.Ribbon;
using Autodesk.Windows;
using System.Windows.Controls;

namespace AsdRcSlab
{
    public static class RibbonBuilder
    {
        public static void Build()
        {
            RibbonControl ribbon = ComponentManager.Ribbon;
            if (ribbon == null) return;

            // Usun istniejaca zakladke jesli juz istnieje
            foreach (RibbonTab existing in ribbon.Tabs)
            {
                if (existing.Id == "ASD_RC_SLAB_TAB")
                {
                    ribbon.Tabs.Remove(existing);
                    break;
                }
            }

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

            ribbon.Tabs.Add(tab);
            tab.IsActive = true;
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
