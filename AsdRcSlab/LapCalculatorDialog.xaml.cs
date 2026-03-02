using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace AsdRcSlab
{
    public partial class LapCalculatorDialog : Window
    {
        private static readonly Dictionary<string, int> LapLengths = new Dictionary<string, int>
        {
            { "H10", 400 },
            { "H12", 500 },
            { "H16", 650 }
        };

        public LapCalculatorDialog()
        {
            InitializeComponent();
            Calculate();
        }

        private void Input_Changed(object sender, object e) => Calculate();

        private void Calculate()
        {
            if (TxtL0 == null) return;

            string dia = (CmbDia?.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "H10";
            int l0 = LapLengths.ContainsKey(dia) ? LapLengths[dia] : 400;

            bool hasSpan = double.TryParse(
                TxtSpan?.Text?.Replace(",", "."),
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double span);

            TxtL0.Text      = $"{l0} mm";
            TxtStagger.Text = $"{(int)(l0 * 0.65)} mm";

            if (hasSpan && span > 0)
            {
                int maxPos = (int)(span * 0.15 * 1000); // mm
                TxtPos.Text = $"≤ {maxPos} mm";

                // Ostrzezenie: pole za krotkie
                double minSpan = l0 * 2.0 / 1000.0;
                if (span < minSpan)
                {
                    TxtWarn.Text        = $"Uwaga: pole ({span:F2} m) jest krótsze niż 2×l₀ = {minSpan:F2} m. " +
                                          $"Zakład może nie zmieścić się poza strefą 15%.";
                    PanelWarn.Visibility = Visibility.Visible;
                }
                else
                {
                    PanelWarn.Visibility = Visibility.Collapsed;
                }
            }
            else
            {
                TxtPos.Text          = "—";
                PanelWarn.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();
    }
}
