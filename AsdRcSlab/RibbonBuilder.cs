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

            tab.Panels.Add(CreatePanel("PROJEKT",
                new[] {
                    ("Nowy Projekt",    "ASD-PROJ", "Tworzy nowy projekt i zapisuje project.json"),
                    ("Otworz Projekt",  "ASD-OPEN", "Wczytuje projekt z pliku project.json"),
                    ("Wczytaj GA",      "ASD-GAI",  "Import geometrii z GA (Sprint 2)"),
                    ("Ustawienia",      "ASD-SET",  "Parametry systemowe Speedeck")
                }));

            tab.Panels.Add(CreatePanel("ZBROJENIE",
                new[] {
                    ("Generuj B1/B2",  "ASD-GBOT", "Auto-generuje siatke dolna H10@200"),
                    ("Generuj T1/T2",  "ASD-GTOP", "Auto-generuje siatke gorna H12@200"),
                    ("Oznacz Prety",   "ASD-BMM",  "Walidacja BBS: R87, R95, R81, R83"),
                    ("Zaklady Auto",   "ASD-LAP",  "Kalkulator zakladow wg Note 16")
                }));

            tab.Panels.Add(CreatePanel("PH CONDITIONS",
                new[] {
                    ("Wczytaj Punching", "ASD-PXIE", "Import PUNCHING_NEW_TEMPLATE_v2.xlsx"),
                    ("Assign PH",        "ASD-PAA",  "Przypisuje PH1-PH9 i generuje tytuly detali"),
                    ("PH Report",        "ASD-PHR",  "Generuje PH_Report.xlsx"),
                    ("Waliduj PH",       "ASD-PHV",  "Sprawdza R77, R79, duplikaty")
                }));

            tab.Panels.Add(CreatePanel("QA VALIDATOR",
                new[] {
                    ("Sprawdz BBS",    "ASD-BBSV", "Waliduje BBS Excel: R87/R95/R81/R83/R92"),
                    ("PIV Check 15R",  "ASD-PIV",  "Uruchamia 15 regul PIV"),
                    ("Raport Bledow",  "ASD-GER",  "Generuje PIV_Dashboard.xlsx"),
                    ("Podglad QA",     "ASD-QAP",  "Status wszystkich regul biezacego projektu")
                }));

            tab.Panels.Add(CreatePanel("EKSPORT",
                new[] {
                    ("Generuj BBS",  "ASD-BSX", "Auto-wypelnia BS8666 BBS.xlsx"),
                    ("PDF Export",   "ASD-PDF", "Auto-PURGE, white check, PDF per arkusz"),
                    ("Calc Doc",     "ASD-CAG", "Wypelnia Template_PiledRaft.docx"),
                    ("Transmittal",  "ASD-TRX", "Generuje transmittal list")
                }));

            ribbon.Tabs.Add(tab);
            tab.IsActive = true;
        }

        private static RibbonPanel CreatePanel(string title, (string label, string cmd, string tooltip)[] buttons)
        {
            RibbonPanelSource source = new RibbonPanelSource { Title = title };
            RibbonRowPanel row = new RibbonRowPanel();

            foreach (var (label, cmd, tooltip) in buttons)
            {
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
                row.Items.Add(new RibbonRowBreak());
            }

            source.Items.Add(row);

            RibbonPanel panel = new RibbonPanel { Source = source };
            return panel;
        }
    }
}
