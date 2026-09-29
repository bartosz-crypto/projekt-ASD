using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace AsdRcSlab
{
    public partial class BbsGeneratorDialog : Window
    {
        public ObservableCollection<BbsLayoutAssignment> AssignmentsCollection { get; set; }

        public List<BbsLayerAssignment> AssignmentOptions { get; }
            = new List<BbsLayerAssignment>
            {
                BbsLayerAssignment.Skip,
                BbsLayerAssignment.Bottom,
                BbsLayerAssignment.Top,
                BbsLayerAssignment.BottomAndTop
            };

        public BbsGenerationContext Result            { get; private set; }
        public string               SelectedOutputPath { get; private set; }

        public BbsGeneratorDialog(BbsGenerationContext initial)
            : this(initial, null) { }

        public BbsGeneratorDialog(
            BbsGenerationContext initial, string suggestedOutputPath)
        {
            InitializeComponent();
            DataContext = this;

            AssignmentsCollection =
                new ObservableCollection<BbsLayoutAssignment>(initial.Assignments);
            LayoutsGrid.ItemsSource = AssignmentsCollection;

            // Output path: pre-fill z suggestion (auto-generated name z .dwg
            // base name + .xls extension w tym samym folderze)
            OutputBox.Text = suggestedOutputPath ?? "";

            ContractNoBox.Text = initial.ContractNo    ?? "";
            Address1Box.Text   = initial.AddressLine1  ?? "";
            Address2Box.Text   = initial.AddressLine2  ?? "";
            Address3Box.Text   = initial.AddressLine3  ?? "";
            RevisionBox.Text   = initial.Revision      ?? "C1";
            PlotSuffixBox.Text = initial.PlotSuffix    ?? "";

            // p166: Accessories
            SelectComboByText(TricTrakTypeBox, initial.TricTrakType, "TT40");
            SelectComboByText(HystoolsTypeBox, initial.HystoolsType, "DK165");
            TricTrakQtyBox.Text = initial.TricTrakQty ?? "";
            HystoolsQtyBox.Text = initial.HystoolsQty ?? "";
            UpdateAccessoryPreview();
        }

        private static void SelectComboByText(
            System.Windows.Controls.ComboBox box, string value, string fallback)
        {
            string want = string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
            foreach (var item in box.Items.OfType<System.Windows.Controls.ComboBoxItem>())
            {
                if (string.Equals(item.Content as string, want,
                                  System.StringComparison.OrdinalIgnoreCase))
                {
                    box.SelectedItem = item;
                    return;
                }
            }
            box.SelectedIndex = 0;
        }

        private static string ComboText(System.Windows.Controls.ComboBox box, string fallback)
        {
            return (box?.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content as string
                   ?? fallback;
        }

        private void OnAccessoryChanged(object sender, RoutedEventArgs e)
        {
            UpdateAccessoryPreview();
        }

        // Podgląd dokładnego tekstu, który trafi do BBS (A42 / I43).
        private void UpdateAccessoryPreview()
        {
            // Handlery odpalają się już w InitializeComponent — kontrolki mogą być null.
            if (TricTrakPreview == null || HystoolsPreview == null ||
                TricTrakQtyBox == null  || HystoolsQtyBox == null) return;

            var tmp = BuildAccessoryContext();
            TricTrakPreview.Text = "→ " + tmp.BuildTricTrakLine();
            HystoolsPreview.Text = "→ " + tmp.BuildHystoolsLine();
        }

        private BbsGenerationContext BuildAccessoryContext()
        {
            return new BbsGenerationContext
            {
                TricTrakType = ComboText(TricTrakTypeBox, "TT40"),
                TricTrakQty  = (TricTrakQtyBox?.Text ?? "").Trim(),
                HystoolsType = ComboText(HystoolsTypeBox, "DK165"),
                HystoolsQty  = (HystoolsQtyBox?.Text ?? "").Trim()
            };
        }

        private static bool IsValidQty(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return true;   // puste = OK (zostaw "No.")
            return int.TryParse(s.Trim(), out int n) && n > 0;
        }

        private void OnOutputBrowseClick(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title           = "Save BBS as...",
                Filter          = "Excel 97-2003 (*.xls)|*.xls|Excel (*.xlsx)|*.xlsx",
                DefaultExt      = ".xls",
                AddExtension    = true,
                OverwritePrompt = false
            };
            if (!string.IsNullOrWhiteSpace(OutputBox.Text))
            {
                var dir = System.IO.Path.GetDirectoryName(OutputBox.Text);
                if (System.IO.Directory.Exists(dir))
                    dlg.InitialDirectory = dir;
                dlg.FileName = System.IO.Path.GetFileName(OutputBox.Text);
            }
            if (dlg.ShowDialog() == true)
                OutputBox.Text = dlg.FileName;
        }

        private void OnOkClick(object sender, RoutedEventArgs e)
        {
            // Walidacja: przynajmniej 1 layout musi mieć assignment ≠ Skip
            bool anyAssigned = AssignmentsCollection
                .Any(a => a.Assignment != BbsLayerAssignment.Skip);
            if (!anyAssigned)
            {
                System.Windows.MessageBox.Show(
                    "Please assign at least one layout to Bottom, Top, or Bottom + Top.\n\n"
                    + "If all layouts are 'Skip', no BBS will be generated.",
                    "No layouts assigned",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;  // NIE zamykaj dialogu
            }

            string output = OutputBox.Text;
            if (string.IsNullOrWhiteSpace(output))
            {
                MessageBox.Show(
                    "Output BBS path is required.",
                    "Missing output",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }
            // Walidacja: folder musi istnieć (plik nie musi — będzie utworzony)
            var dir2 = System.IO.Path.GetDirectoryName(output);
            if (!string.IsNullOrEmpty(dir2) && !System.IO.Directory.Exists(dir2))
            {
                MessageBox.Show(
                    "Output folder doesn't exist:\n" + dir2,
                    "Invalid folder",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // p166: ilości accessories — puste albo dodatnia liczba całkowita
            if (!IsValidQty(TricTrakQtyBox.Text) || !IsValidQty(HystoolsQtyBox.Text))
            {
                MessageBox.Show(
                    "TRIC-TRAK / HYSTOOLS quantity must be a whole number (e.g. 52) or empty.",
                    "Invalid quantity",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }
            var acc = BuildAccessoryContext();

            Result = new BbsGenerationContext
            {
                Assignments  = new List<BbsLayoutAssignment>(AssignmentsCollection),
                ContractNo   = ContractNoBox.Text,
                AddressLine1 = Address1Box.Text,
                AddressLine2 = Address2Box.Text,
                AddressLine3 = Address3Box.Text,
                Revision     = RevisionBox.Text,
                PlotSuffix   = PlotSuffixBox.Text,
                TricTrakType = acc.TricTrakType,
                TricTrakQty  = acc.TricTrakQty,
                HystoolsType = acc.HystoolsType,
                HystoolsQty  = acc.HystoolsQty
            };

            SelectedOutputPath = output;

            DialogResult = true;
            Close();
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
