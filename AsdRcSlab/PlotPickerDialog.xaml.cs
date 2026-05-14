using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace AsdRcSlab
{
    public partial class PlotPickerDialog : Window
    {
        public PlotInfo SelectedPlot { get; private set; }

        private readonly List<PlotDisplayItem> _items;

        public PlotPickerDialog(List<PlotInfo> plots)
        {
            InitializeComponent();
            _items = plots.Select(p => new PlotDisplayItem(p)).ToList();
            PlotsList.ItemsSource = _items;
            if (_items.Count > 0) PlotsList.SelectedIndex = 0;
        }

        private void OkClick(object sender, RoutedEventArgs e)
        {
            var item = PlotsList.SelectedItem as PlotDisplayItem;
            if (item == null) return;
            SelectedPlot = item.Info;
            DialogResult = true;
        }

        private void CancelClick(object sender, RoutedEventArgs e) => DialogResult = false;

        private void PlotsList_DoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (PlotsList.SelectedItem != null) OkClick(sender, e);
        }

        private class PlotDisplayItem
        {
            public PlotInfo Info  { get; }
            public string   Line1 { get; }
            public string   Line2 { get; }

            public PlotDisplayItem(PlotInfo info)
            {
                Info  = info;
                Line1 = $"PLOT {info.Number} ({info.PileCount} piles)";

                var sb = new StringBuilder($"INT: {info.InternalCount}  EDGE: {info.EdgeCount}  CORNER: {info.CornerCount}");
                if (info.ReentrantCount > 0)
                    sb.Append($"  REENTRANT: {info.ReentrantCount}");
                Line2 = sb.ToString();
            }
        }
    }
}
