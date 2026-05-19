using System.Collections.Generic;
using System.Windows;

namespace AsdRcSlab
{
    public partial class ImrPlotPickerDialog : Window
    {
        public ImrCommand.PlotMapInfo SelectedPlot { get; private set; }

        private readonly List<ImrCommand.PlotMapInfo> _plots;

        public ImrPlotPickerDialog(List<ImrCommand.PlotMapInfo> plots)
        {
            InitializeComponent();
            _plots = plots;
            foreach (var p in _plots)
                PlotList.Items.Add(p.Label);
            if (PlotList.Items.Count > 0)
                PlotList.SelectedIndex = 0;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            int idx = PlotList.SelectedIndex;
            if (idx < 0)
            {
                MessageBox.Show("Please select a plot.", "No selection",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            SelectedPlot = _plots[idx];
            DialogResult = true;
        }

        private void PlotList_DoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (PlotList.SelectedItem != null) OkButton_Click(sender, e);
        }
    }
}
