using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace AsdRcSlab
{
    public partial class SheetPickerDialog : Window
    {
        public string SelectedSheet { get; private set; }

        public SheetPickerDialog(List<string> sheets)
        {
            InitializeComponent();
            LstSheets.ItemsSource = sheets;
            if (sheets.Count > 0)
                LstSheets.SelectedIndex = 0;
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            if (LstSheets.SelectedItem == null)
            {
                MessageBox.Show("Wybierz arkusz z listy.", "Brak wyboru",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            SelectedSheet = LstSheets.SelectedItem as string;
            DialogResult = true;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;

        private void LstSheets_DoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (LstSheets.SelectedItem != null)
                BtnOk_Click(sender, e);
        }
    }
}
