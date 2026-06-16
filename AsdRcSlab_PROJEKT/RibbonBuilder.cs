using Autodesk.AutoCAD.Ribbon;
using Autodesk.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace AsdRcSlab
{
    public static class RibbonBuilder
    {
        // Mapowanie komenda -> plik ikony (w folderze Contents\Icons obok DLL).
        // Aby zmienic ikone przycisku wystarczy podmienic nazwe pliku ponizej.
        private static readonly System.Collections.Generic.Dictionary<string, string> IconMap =
            new System.Collections.Generic.Dictionary<string, string>
        {
            { "ASD-GAI",       "Pino-Disney-Mickey-Mouse-1.32.png" },   // Copy from GA
            { "ASD-RCN",       "Pino-Disney-Minnie-Mouse.32.png"   },   // Sheet Numbering
            { "ASD-PXIE",      "SD-EditPile.32.png"                },   // Load Punching
            { "ASD-PAA",       "SD-MarkPiles.32.png"               },   // Assign PH
            { "ASD-PHR",       "SD-MasterLevels.32.png"            },   // PH Report
            { "ASD-PHV",       "Iconfactory-Looney-Y.-Sam.32.png"  },   // Waliduj PH
            { "ASD-IMR",       "Iconfactory-Looney-Roadrunner.32.png" },// Import Maps
            { "ASD-BBC",       "Iconfactory-Looney-Taz.32.png"     },   // Bar Calculator
            { "ASD-BBS-WRITE", "Iconfactory-Looney-Tweety.32.png"  },   // BBS Write
            // ASD-ABOUT (NA Engineering) celowo bez ikony - sam tekst + link.
        };

        // Wersja dodatku + data builda (z czasu modyfikacji DLL = czas kompilacji).
        internal const string Version = "4.4";

        internal static string BuildStamp()
        {
            try
            {
                var loc = System.Reflection.Assembly.GetExecutingAssembly().Location;
                return System.IO.File.GetLastWriteTime(loc).ToString("yyyy-MM-dd HH:mm");
            }
            catch { return "?"; }
        }

        private static string IconsDir()
        {
            var dll = System.Reflection.Assembly.GetExecutingAssembly().Location;
            return System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(dll) ?? "", "Icons");
        }

        // Pelna sciezka do pliku ikony dla danej komendy (null jesli brak mapowania/pliku).
        internal static string IconPath(string cmd)
        {
            if (cmd == null || !IconMap.TryGetValue(cmd, out var file)) return null;
            var path = System.IO.Path.Combine(IconsDir(), file);
            if (!System.IO.File.Exists(path))
            {
                App.DiagLog($"[Icon] missing: {path}");
                return null;
            }
            return path;
        }

        // Wczytuje PNG z dysku do zamrozonego ImageSource (bezpieczne miedzy watkami).
        private static ImageSource LoadIcon(string cmd)
        {
            var path = IconPath(cmd);
            if (path == null) return null;
            try
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = new System.Uri(path, System.UriKind.Absolute);
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
            catch (System.Exception ex)
            {
                App.DiagLog($"[Icon] load err {cmd}: {ex.Message}");
                return null;
            }
        }

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

            // p128/p131: panel narzedzi (tekstowo, bez ikon).
            tab.Panels.Add(CreatePanel("TOOLS",
                new[] {
                    ("Explode ASD", "ASD-XAS",
                     "Explode AutoCAD Structural Detailing objects "
                     + "(layers 'AutoCAD_Structural_Detailing_*') into plain geometry. "
                     + "Whole model space; originals deleted."),
                    ("Scale Detail Circles", "ASD-SDC",
                     "Detect 1:25 detail frames (dashed color 1/10) containing "
                     + "distribution circles and scale those circles by 0.5 "
                     + "(preview dialog; default 1:50 R=37.5 -> 18.75).")
                }, columnsPerRow: 1));

            // Panel wlasciciela + wersja/data builda (tekstowo, bez ikon).
            {
                RibbonPanelSource src = new RibbonPanelSource { Title = "NA ENGINEERING" };
                RibbonRowPanel prow = new RibbonRowPanel();

                prow.Items.Add(new RibbonButton
                {
                    Text = "NA Engineering sp. z o.o.",
                    CommandHandler = new RibbonCommandHandler("ASD-ABOUT"),
                    CommandParameter = "ASD-ABOUT",
                    ShowText = true,
                    ShowImage = false,
                    Size = RibbonItemSize.Standard,
                    Width = 180,
                    MinWidth = 180,
                    ToolTip = "Wlasciciel dodatku - otworz https://naengineering.uk"
                });
                prow.Items.Add(new RibbonRowBreak());
                prow.Items.Add(new RibbonLabel
                {
                    Text = $"v{Version}  build {BuildStamp()}",
                    Width = 180,
                    ToolTip = "Wersja dodatku i data/godzina kompilacji aktualnego buildu"
                });

                src.Items.Add(prow);
                tab.Panels.Add(new RibbonPanel { Source = src });
            }

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

            // === p125 DIAG: workspace info ===
            try
            {
                // Aktualny workspace name
                object wsName = null;
                try { wsName = Autodesk.AutoCAD.ApplicationServices.Application
                    .GetSystemVariable("WSCURRENT"); } catch { }
                App.DiagLog($"[Build] WSCURRENT = {wsName}");

                // Ile tabow widzi ComponentManager
                App.DiagLog($"[Build] ComponentManager.Ribbon.Tabs.Count = {ComponentManager.Ribbon.Tabs.Count}");

                // Czy nasz tab jest w kolekcji (potwierdzenie)
                bool found = false;
                foreach (Autodesk.Windows.RibbonTab t in ComponentManager.Ribbon.Tabs)
                    if (t.Id == "ASD_RC_SLAB_TAB") { found = true; break; }
                App.DiagLog($"[Build] nasz tab w kolekcji = {found}");
            }
            catch (System.Exception ex)
            {
                App.DiagLog($"[Build] workspace DIAG EXCEPTION: {ex.Message}");
            }

            // === p125 Podejście A: reload current workspace ===
            // ASD 2015 renderuje ribbon z workspace - przeladowanie biezacego
            // workspace wymusza przebudowe ribbon z aktualnej kolekcji
            // ComponentManager (wlacznie z naszym tab).
            try
            {
                var wsName = Autodesk.AutoCAD.ApplicationServices.Application
                    .GetSystemVariable("WSCURRENT") as string;
                if (!string.IsNullOrEmpty(wsName))
                {
                    var doc = Autodesk.AutoCAD.ApplicationServices.Application
                        .DocumentManager.MdiActiveDocument;
                    if (doc != null)
                    {
                        // Reaktywacja workspace przez komende WSCURRENT.
                        // Escaping cudzyslowow dla nazw z spacjami.
                        string cmd = $"_.WSCURRENT \"{wsName}\"\n";
                        doc.SendStringToExecute(cmd, true, false, false);
                        App.DiagLog($"[Build] WSCURRENT reload wyslany dla '{wsName}'");
                    }
                    else
                    {
                        App.DiagLog("[Build] brak active document - nie mozna reload workspace");
                    }
                }
            }
            catch (System.Exception ex)
            {
                App.DiagLog($"[Build] workspace reload EXCEPTION: {ex.Message}");
            }

            // === p126b DIAG: sondowanie workspace API przez reflection ===
            try
            {
                // 1. Czy istnieje CustomizationSection API?
                var acMgdAsm = System.Reflection.Assembly.GetAssembly(
                    typeof(Autodesk.AutoCAD.ApplicationServices.Application));
                App.DiagLog($"[API] AcMgd assembly: {acMgdAsm?.GetName().Version}");

                // 2. Szukaj typow zwiazanych z workspace/customization
                var allAsms = System.AppDomain.CurrentDomain.GetAssemblies();
                foreach (var asm in allAsms)
                {
                    var name = asm.GetName().Name;
                    if (name == "AcMgd" || name == "AcCoreMgd" || name == "AcCui" || name == "AcWindows")
                    {
                        App.DiagLog($"[API] Assembly loaded: {name} v{asm.GetName().Version}");
                        try
                        {
                            foreach (var t in asm.GetExportedTypes())
                            {
                                var tn = t.FullName ?? "";
                                if (tn.Contains("Workspace") || tn.Contains("Customization") ||
                                    tn.Contains("Cui") || tn.Contains("Ribbon"))
                                {
                                    App.DiagLog($"[API]   type: {tn}");
                                }
                            }
                        }
                        catch (System.Exception ex) { App.DiagLog($"[API]   GetExportedTypes err: {ex.Message}"); }
                    }
                }
            }
            catch (System.Exception ex)
            {
                App.DiagLog($"[API] DIAG EXCEPTION: {ex.Message}");
            }

            // === p126b DIAG: RibbonTab properties ===
            try
            {
                var tabType = tab.GetType();
                App.DiagLog($"[TAB] type: {tabType.FullName}");
                foreach (var prop in tabType.GetProperties())
                {
                    var pn = prop.Name;
                    // Loguj property zwiazane z identyfikacja / workspace / visibility
                    if (pn.Contains("Id") || pn.Contains("Uid") || pn.Contains("Workspace") ||
                        pn.Contains("Visible") || pn.Contains("Anonymous") || pn.Contains("Tag") ||
                        pn.Contains("Name"))
                    {
                        object val = null;
                        try { val = prop.GetValue(tab); } catch { }
                        App.DiagLog($"[TAB]   {pn} = {val}");
                    }
                }
            }
            catch (System.Exception ex)
            {
                App.DiagLog($"[TAB] DIAG EXCEPTION: {ex.Message}");
            }
        }

        private static RibbonPanel CreatePanel(
            string title,
            (string label, string cmd, string tooltip)[] buttons,
            int columnsPerRow = 1)
        {
            // Uklad sprawdzony i dzialajacy: male przyciski (Standard) w RibbonRowPanel.
            // Ikona w slocie 16x16 (Image). Wersja z Size=Large nie renderowala ikon
            // w ASD 2015 - dlatego wracamy do tego ukladu.
            RibbonPanelSource source = new RibbonPanelSource { Title = title };
            RibbonRowPanel row = new RibbonRowPanel();

            for (int i = 0; i < buttons.Length; i++)
            {
                var (label, cmd, tooltip) = buttons[i];
                var icon = LoadIcon(cmd);
                RibbonButton btn = new RibbonButton
                {
                    Text = label,
                    CommandHandler = new RibbonCommandHandler(cmd),
                    CommandParameter = cmd,
                    ShowText = true,
                    ShowImage = icon != null,
                    Image = icon,        // slot 16x16
                    LargeImage = icon,   // slot 32x32 (gdyby panel przeszedl w tryb Large)
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
