using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace AsdRcSlab
{
    // Wiersz listy warstw ASD w dialogu (checkbox + nazwa + licznik).
    public class LayerSelectItem
    {
        public bool Selected { get; set; } = true;
        public string LayerName { get; set; }
        public int Count { get; set; }
    }

    public partial class ExplodeAsdDialog : Window
    {
        private readonly List<LayerSelectItem> _items;

        public bool Confirmed { get; private set; }
        public bool Recursive { get; private set; } = true;
        public bool RecolorDistributionCircles { get; private set; } = true;
        public List<string> SelectedLayers { get; private set; } = new List<string>();

        public ExplodeAsdDialog(List<(string Layer, int Count)> layers)
        {
            InitializeComponent();

            _items = (layers ?? new List<(string, int)>())
                .Select(l => new LayerSelectItem
                {
                    Selected = true,
                    LayerName = l.Layer,
                    Count = l.Count
                })
                .ToList();

            Grid.ItemsSource = _items;

            int totalObjects = _items.Sum(i => i.Count);
            TxtInfo.Text =
                $"{_items.Count} ASD layer(s), {totalObjects} top-level object(s) " +
                $"found in model space.";
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var i in _items) i.Selected = true;
            Grid.Items.Refresh();
        }

        private void BtnSelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var i in _items) i.Selected = false;
            Grid.Items.Refresh();
        }

        private void BtnExplode_Click(object sender, RoutedEventArgs e)
        {
            // commit dowolnej aktywnej edycji w gridzie
            Grid.CommitEdit();

            SelectedLayers = _items.Where(i => i.Selected)
                                   .Select(i => i.LayerName)
                                   .ToList();

            if (SelectedLayers.Count == 0)
            {
                MessageBox.Show(this,
                    "Select at least one ASD layer to explode.",
                    "Explode ASD Objects",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Recursive = ChkRecursive.IsChecked == true;
            RecolorDistributionCircles = ChkRecolorCircles.IsChecked == true;
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
