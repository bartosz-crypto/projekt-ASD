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

            tab.Panels.Add(CreatePanel("TABELKA TYTULOWA",
                new[] {
                    ("Opisz z GA",      "ASD-GAI", "Import atrybutow z GA do tabelek tytulowych RC"),
                    ("Nazwy Rysunkow",  "ASD-RCN", "Auto-wypelnienie TITLE_3, SCALE, DATE + rename layoutow")
                }, columnsPerRow: 2));

            tab.Panels.Add(CreatePanel("PH CONDITIONS",
                new[] {
                    ("Wczytaj Punching", "ASD-PXIE", "Import PUNCHING_NEW_TEMPLATE_v2.xlsx"),
                    ("Assign PH",        "ASD-PAA",  "Przypisuje PH1-PH9 i generuje tytuly detali"),
                    ("PH Report",        "ASD-PHR",  "Generuje PH_Report.xlsx"),
                    ("Waliduj PH",       "ASD-PHV",  "Sprawdza R77, R79, duplikaty")
                }, columnsPerRow: 2));

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
                    Orientation = Orientation.Vertical,
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
