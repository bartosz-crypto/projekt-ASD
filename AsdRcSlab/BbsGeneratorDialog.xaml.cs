using System.Collections.Generic;
using System.Collections.ObjectModel;
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
                BbsLayerAssignment.Top
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

            Result = new BbsGenerationContext
            {
                Assignments  = new List<BbsLayoutAssignment>(AssignmentsCollection),
                ContractNo   = ContractNoBox.Text,
                AddressLine1 = Address1Box.Text,
                AddressLine2 = Address2Box.Text,
                AddressLine3 = Address3Box.Text,
                Revision     = RevisionBox.Text,
                PlotSuffix   = PlotSuffixBox.Text
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
