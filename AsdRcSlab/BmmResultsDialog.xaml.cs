using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;

namespace AsdRcSlab
{
    public class BmmResultViewModel
    {
        public string RuleId  { get; set; }
        public string Status  { get; set; }
        public string Details { get; set; }
        public Brush  Background { get; set; }
    }

    public partial class BmmResultsDialog : Window
    {
        private readonly BmmResult _result;

        public BmmResultsDialog(BmmResult result, string fileName)
        {
            InitializeComponent();
            _result = result;
            Title   = $"BMM — {fileName}";
            PopulateResults();
        }

        private void PopulateResults()
        {
            var rules = new[]
            {
                _result.R87, _result.R95, _result.R81, _result.R83, _result.R92
            };

            var items = rules.Select(r => new BmmResultViewModel
            {
                RuleId     = r.RuleId,
                Status     = r.Status,
                Details    = r.Details,
                Background = r.Status == "FAIL" ? new SolidColorBrush(Color.FromRgb(0xFC, 0xE4, 0xEC))
                           : r.Status == "WARN" ? new SolidColorBrush(Color.FromRgb(0xFF, 0xF8, 0xDC))
                                                : new SolidColorBrush(Color.FromRgb(0xE2, 0xEF, 0xDA))
            }).ToList();

            LstResults.ItemsSource = items;

            int fails  = rules.Count(r => r.Status == "FAIL");
            int warns  = rules.Count(r => r.Status == "WARN");
            int ok     = rules.Count(r => r.Status == "OK");

            TxtSummary.Text = $"FAIL: {fails}   WARN: {warns}   OK: {ok}";

            if (fails > 0)
            {
                SummaryBorder.Background = new SolidColorBrush(Color.FromRgb(0xFC, 0xE4, 0xEC));
                TxtSummary.Foreground    = new SolidColorBrush(Color.FromRgb(0xB7, 0x1C, 0x1C));
            }
            else if (warns > 0)
            {
                SummaryBorder.Background = new SolidColorBrush(Color.FromRgb(0xFF, 0xF8, 0xDC));
                TxtSummary.Foreground    = new SolidColorBrush(Color.FromRgb(0xF5, 0x7F, 0x17));
            }
            else
            {
                SummaryBorder.Background = new SolidColorBrush(Color.FromRgb(0xE2, 0xEF, 0xDA));
                TxtSummary.Foreground    = new SolidColorBrush(Color.FromRgb(0x1B, 0x5E, 0x20));
            }
        }

        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            var sb = new StringBuilder();
            sb.AppendLine("BMM — WALIDACJA BBS");
            sb.AppendLine(new string('-', 40));
            foreach (var r in new[] { _result.R87, _result.R95, _result.R81, _result.R83, _result.R92 })
                sb.AppendLine($"{r.RuleId}: {r.Status} — {r.Details}");
            Clipboard.SetText(sb.ToString());
            MessageBox.Show("Raport skopiowany do schowka.", "BMM",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();
    }
}
