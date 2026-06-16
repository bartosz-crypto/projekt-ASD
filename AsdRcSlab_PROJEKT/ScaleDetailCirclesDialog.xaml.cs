using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace AsdRcSlab
{
    // Wiersz listy kandydatów (checkbox + opis ramki).
    public class ScaleDetailRow
    {
        public bool Selected { get; set; }
        public string Detail { get; set; }
        public string Scale { get; set; }
        public int Circles { get; set; }
        public string Center { get; set; }
        public string Size { get; set; }
        public string Dist { get; set; }
        public bool Uncertain { get; set; }
        public DetailCandidate Candidate { get; set; }
    }

    public partial class ScaleDetailCirclesDialog : Window
    {
        private readonly List<ScaleDetailRow> _rows;

        public bool Confirmed { get; private set; }
        public List<DetailCandidate> SelectedCandidates { get; private set; }
            = new List<DetailCandidate>();

        public ScaleDetailCirclesDialog(List<DetailCandidate> candidates)
        {
            InitializeComponent();

            _rows = (candidates ?? new List<DetailCandidate>())
                .OrderByDescending(c => c.Preselected)
                .ThenBy(c => c.NearestLabelDist)
                .Select(c => new ScaleDetailRow
                {
                    Selected = c.Preselected,
                    Detail = c.NearestLabelText,
                    Scale = "1:25",
                    Circles = c.CircleCount,
                    Center = $"({c.CenterX:F0}, {c.CenterY:F0})",
                    Size = $"{c.W:F0} × {c.H:F0}",
                    Dist = c.Preselected
                        ? $"{c.NearestLabelDist:F0}"
                        : $"{c.NearestLabelDist:F0} (check!)",
                    Uncertain = !c.Preselected,
                    Candidate = c
                })
                .ToList();

            Grid.ItemsSource = _rows;

            int circlesTotal = _rows.Sum(r => r.Circles);
            int preselected = _rows.Count(r => r.Selected);
            TxtInfo.Text =
                $"{_rows.Count} candidate frame(s), {circlesTotal} circle(s). " +
                $"{preselected} pre-selected (near a 1:25 label). " +
                $"Unchecked rows marked '(check!)' — verify before Apply. " +
                $"Apply scales checked frames' circles ×0.5.";
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in _rows) r.Selected = true;
            Grid.Items.Refresh();
        }

        private void BtnSelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in _rows) r.Selected = false;
            Grid.Items.Refresh();
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            Grid.CommitEdit();
            SelectedCandidates = _rows.Where(r => r.Selected)
                                      .Select(r => r.Candidate)
                                      .ToList();
            if (SelectedCandidates.Count == 0)
            {
                MessageBox.Show(this,
                    "Check at least one detail frame to scale.",
                    "Scale Detail Circles",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Confirmed = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            Confirmed = false;
            Close();
        }
    }
}
