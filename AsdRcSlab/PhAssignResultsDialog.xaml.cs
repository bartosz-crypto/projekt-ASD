using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace AsdRcSlab
{
    public class PileViewModel
    {
        public string PileId       { get; set; }
        public string UtilPctStr   { get; set; }
        public string LocationType { get; set; }
        public string PhAction     { get; set; }
        public string DetailTitle  { get; set; }
    }

    public partial class PhAssignResultsDialog : Window
    {
        private readonly List<PileData> _piles;

        public PhAssignResultsDialog(List<PileData> piles)
        {
            InitializeComponent();
            _piles = piles;
            Populate();
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

            var counts = _piles.GroupBy(p => p.PhAction)
                .OrderBy(g => g.Key)
                .Select(g => $"{g.Key}: {g.Count()}")
                .ToArray();

            int exceed = _piles.Count(p => p.PhAction == "EXCEED");
            string exceedWarn = exceed > 0 ? $"  ⚠ EXCEED: {exceed}" : "";
            TxtStats.Text = $"Razem: {_piles.Count} pali   |   " +
                            string.Join("  ", counts) + exceedWarn;
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

                    // Naglowki
                    string[] headers = { "Pile ID", "Util %", "Location", "PH Action", "Tytuł Detalu" };
                    for (int c = 0; c < headers.Length; c++)
                    {
                        ws.Cells[1, c + 1].Value = headers[c];
                        ws.Cells[1, c + 1].Style.Font.Bold = true;
                        ws.Cells[1, c + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[1, c + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x15, 0x65, 0xC0));
                        ws.Cells[1, c + 1].Style.Font.Color.SetColor(Color.White);
                    }

                    // Dane
                    for (int i = 0; i < _piles.Count; i++)
                    {
                        var p = _piles[i];
                        int row = i + 2;
                        ws.Cells[row, 1].Value = p.PileId;
                        ws.Cells[row, 2].Value = $"{p.UtilPct:F1}%";
                        ws.Cells[row, 3].Value = p.LocationType;
                        ws.Cells[row, 4].Value = p.PhAction;
                        ws.Cells[row, 5].Value = p.DetailTitle;

                        // Kolorowanie wg PH
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

        private void BtnCopyTitles_Click(object sender, RoutedEventArgs e)
        {
            var sb = new StringBuilder();
            foreach (var p in _piles.OrderBy(x => x.PhAction))
                sb.AppendLine(p.DetailTitle);
            Clipboard.SetText(sb.ToString());
            MessageBox.Show("Tytuły detali skopiowane do schowka.", "Kopiuj tytuły",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

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
