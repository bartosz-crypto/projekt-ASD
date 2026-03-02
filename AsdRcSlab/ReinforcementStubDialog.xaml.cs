using System;
using System.Windows;

namespace AsdRcSlab
{
    public partial class ReinforcementStubDialog : Window
    {
        private readonly int _dia;        // 10 lub 12
        private readonly string _layer;   // "B1/B2" lub "T1/T2"
        private readonly int _asBase;     // 393 lub 565 mm²/m

        public ReinforcementStubDialog(int dia, string layer, int asBase)
        {
            InitializeComponent();
            _dia    = dia;
            _layer  = layer;
            _asBase = asBase;

            TxtHeader.Text   = $"GENERUJ {layer} — H{dia}@200";
            TxtMeshInfo.Text = $"H{dia}@200  |  As = {asBase} mm²/m";

            // Wstaw powierzchnię z projektu jeśli wczytany
            if (SessionData.CurrentProject != null)
                TxtArea.Text = "0";
        }

        private void BtnSimulate_Click(object sender, RoutedEventArgs e)
        {
            if (!double.TryParse(TxtArea.Text.Replace(",", "."),
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double area) || area <= 0)
            {
                MessageBox.Show("Podaj poprawną powierzchnię płyty (m²).",
                    "Błąd", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Wzór: area / 0.2 prętów per kierunek, 2 kierunki
            int barsPerDir = (int)Math.Ceiling(area / 0.2);
            int totalBars  = barsPerDir * 2;

            // Masa szacunkowa: długość ~ sqrt(area), 2 kierunki
            double side   = Math.Sqrt(area);
            double length = side * barsPerDir * 2; // metry łączne
            double unitMass = Math.PI * Math.Pow(_dia / 2000.0, 2) * 7850; // kg/m
            double totalMass = length * unitMass;

            TxtResult.Text =
                $"Siatka {_layer}  —  H{_dia}@200\n" +
                $"Pręty na kierunek:  {barsPerDir} szt.\n" +
                $"Pręty łącznie:      {totalBars} szt. (2 kierunki)\n" +
                $"Długość łączna:     {length:F1} m\n" +
                $"Masa szacunkowa:    {totalMass:F0} kg\n\n" +
                $"Pełne rysowanie w ASD → Sprint 2";

            PanelResult.Visibility = Visibility.Visible;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();
    }
}
