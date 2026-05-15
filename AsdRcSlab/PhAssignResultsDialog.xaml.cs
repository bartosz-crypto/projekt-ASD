using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;

namespace AsdRcSlab
{
    public class PileViewModel : INotifyPropertyChanged
    {
        public string PileId       { get; set; }
        public string UtilPctStr   { get; set; }
        public string LocationType { get; set; }
        public string DetailTitle  { get; set; }

        private string _phAction;
        public string PhAction
        {
            get => _phAction;
            set
            {
                if (_phAction == value) return;
                _phAction = value;
                var pile = SessionData.Piles?.FirstOrDefault(p =>
                    string.Equals(p.PileId, PileId, StringComparison.OrdinalIgnoreCase));
                if (pile != null)
                    pile.PhAction = value;
                OnPropertyChanged();
                PhActionChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public static event EventHandler PhActionChanged;

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class PhAssignResultsDialog : Window
    {
        private readonly List<PileData> _piles;

        public PhAssignResultsDialog(List<PileData> piles, bool showUpdateButton = false)
        {
            InitializeComponent();
            _piles = piles;

            BtnUpdateDrawing.Visibility = showUpdateButton
                ? Visibility.Visible
                : Visibility.Collapsed;

            Grid.IsReadOnly = !showUpdateButton;

            Populate();

            if (showUpdateButton)
            {
                PileViewModel.PhActionChanged += OnPhActionChanged;
                this.Closed += (s, e) => { PileViewModel.PhActionChanged -= OnPhActionChanged; };
            }
        }

        private void OnPhActionChanged(object sender, EventArgs e)
        {
            UpdateStats();
            Grid.Items.Refresh();
        }

        private void Populate()
        {
            var vms = _piles.Select(p => new PileViewModel
            {
                PileId       = p.PileId,
                UtilPctStr   = $"{p.UtilPct:F1}%",
                LocationType = p.LocationType,
                PhAction     = p.PhAction,
                DetailTitle  = p.DetailTitle
            }).ToList();

            Grid.ItemsSource = vms;
            UpdateStats();
        }

        private void UpdateStats()
        {
            int total = _piles.Count;

            int CountFor(string ph) => _piles.Count(p =>
                string.Equals(p.PhAction, ph, StringComparison.OrdinalIgnoreCase));

            int p1 = CountFor("PH1"), p2 = CountFor("PH2"), p3 = CountFor("PH3");
            int p3re = CountFor("PH3-RE");
            int p4 = CountFor("PH4"), p5 = CountFor("PH5"), p6 = CountFor("PH6");
            int p7 = CountFor("PH7"), p8 = CountFor("PH8"), p9 = CountFor("PH9");
            int exceed = CountFor("EXCEED"), noAct = CountFor("NO ACTION");

            TxtStats.Text =
                $"Razem: {total} pali  |  " +
                $"PH1:{p1}  PH2:{p2}  PH3:{p3}  PH3-RE:{p3re}  " +
                $"PH4:{p4}  PH5:{p5}  PH6:{p6}  PH7:{p7}  PH8:{p8}  PH9:{p9}  " +
                $"EXCEED:{exceed}  NO ACTION:{noAct}";

            TxtTotals.Text = BuildPhTotalsLine(_piles);
        }

        // Zwraca sformatowany TOTAL line dla statystyk PH.
        // Używany w UpdateStats() oraz w Commands.CmdAssignPH MessageBox.
        internal static string BuildPhTotalsLine(IEnumerable<PileData> piles)
        {
            if (piles == null) return "";

            int CountFor(string ph) => piles.Count(p =>
                string.Equals(p.PhAction, ph, StringComparison.OrdinalIgnoreCase));

            int p1 = CountFor("PH1"), p2 = CountFor("PH2"), p3 = CountFor("PH3");
            int p4 = CountFor("PH4"), p5 = CountFor("PH5"), p6 = CountFor("PH6");
            int p7 = CountFor("PH7"), p8 = CountFor("PH8"), p9 = CountFor("PH9");

            int sumH12  = p1 + p2 + p3;
            int sumH16a = p4 + p5 + p6;
            int sumH16b = p7 + p8 + p9;

            string h12;
            if (sumH12 == 0) h12 = "H12: —";
            else h12 = $"H12: {sumH12}×14 = {sumH12 * 14}";

            string h16;
            if (sumH16a + sumH16b == 0) h16 = "H16: —";
            else if (sumH16b == 0) h16 = $"H16: {sumH16a}×14 = {sumH16a * 14}";
            else if (sumH16a == 0) h16 = $"H16: {sumH16b}×28 = {sumH16b * 28}";
            else h16 = $"H16: {sumH16a}×14+{sumH16b}×28 = {sumH16a * 14 + sumH16b * 28}";

            return $"TOTAL:  {h12}  |  {h16}";
        }

        private void BtnUpdateDrawing_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var res = DrawingAnnotator.Annotate(SessionData.Piles);

                if (res.WrongDrawing)
                {
                    MessageBox.Show(
                        "Aktywny rysunek nie wygląda jak RC SLAB " +
                        "(brak nagłówka 'REINFORCEMENT DETAILS OF SPEEDECK').\n\n" +
                        "Otwórz właściwy rysunek RC i spróbuj ponownie.",
                        "Zaktualizuj rysunek",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string totalsLine = BuildPhTotalsLine(SessionData.Piles);
                string msg =
                    $"Zaktualizowano rysunek.\n\n" +
                    $"Podpisano pali: {res.Annotated.Count}\n" +
                    $"Pominięto (NO ACTION): {res.Skipped.Count}\n" +
                    $"Nie znaleziono: {res.NotFound.Count}\n" +
                    $"Szablony PH (AP-TEXT): {res.PhLabelsUpdated}\n\n" +
                    totalsLine;
                MessageBox.Show(msg, "Zaktualizuj rysunek",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Błąd podczas aktualizacji rysunku:\n{ex.Message}",
                    "Zaktualizuj rysunek",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var saveDlg = new Microsoft.Win32.SaveFileDialog
            {
                Title      = "Zapisz PH Report",
                Filter     = "Excel (*.xlsx)|*.xlsx",
                FileName   = $"PH_Report_{SessionData.CurrentProject?.DRWNumber ?? "export"}_{DateTime.Today:yyyyMMdd}.xlsx"
            };
            if (saveDlg.ShowDialog() != true) return;

            try
            {
                using (var pkg = new ExcelPackage())
                {
                    var ws = pkg.Workbook.Worksheets.Add("PH REPORT");

                    string[] headers = { "Pile ID", "Util %", "Location", "PH Action", "Tytuł Detalu" };
                    for (int c = 0; c < headers.Length; c++)
                    {
                        ws.Cells[1, c + 1].Value = headers[c];
                        ws.Cells[1, c + 1].Style.Font.Bold = true;
                        ws.Cells[1, c + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[1, c + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x15, 0x65, 0xC0));
                        ws.Cells[1, c + 1].Style.Font.Color.SetColor(Color.White);
                    }

                    for (int i = 0; i < _piles.Count; i++)
                    {
                        var p = _piles[i];
                        int row = i + 2;
                        ws.Cells[row, 1].Value = p.PileId;
                        ws.Cells[row, 2].Value = $"{p.UtilPct:F1}%";
                        ws.Cells[row, 3].Value = p.LocationType;
                        ws.Cells[row, 4].Value = p.PhAction;
                        ws.Cells[row, 5].Value = p.DetailTitle;

                        Color bg = GetPhColor(p.PhAction);
                        for (int c = 1; c <= 5; c++)
                        {
                            ws.Cells[row, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, c].Style.Fill.BackgroundColor.SetColor(bg);
                        }
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    pkg.SaveAs(new FileInfo(saveDlg.FileName));
                }

                MessageBox.Show($"Zapisano: {saveDlg.FileName}", "PH Report",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Błąd zapisu: {ex.Message}", "Błąd",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static Color GetPhColor(string ph)
        {
            switch (ph)
            {
                case "PH1":    return Color.FromArgb(0xE2, 0xEF, 0xDA);
                case "PH3-RE": return Color.FromArgb(0xF3, 0xE5, 0xF5);
                case "PH7": case "PH8": case "PH9":
                    return Color.FromArgb(0xFC, 0xE4, 0xEC);
                case "EXCEED": return Color.FromArgb(0xB7, 0x1C, 0x1C);
                default:       return Color.FromArgb(0xFF, 0xF8, 0xDC);
            }
        }
    }
}
