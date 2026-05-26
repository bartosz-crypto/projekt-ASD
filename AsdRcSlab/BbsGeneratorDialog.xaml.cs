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

        public BbsGenerationContext Result           { get; private set; }
        public string               SelectedTemplatePath { get; private set; }

        public BbsGeneratorDialog(BbsGenerationContext initial)
        {
            InitializeComponent();
            DataContext = this;

            AssignmentsCollection =
                new ObservableCollection<BbsLayoutAssignment>(initial.Assignments);
            LayoutsGrid.ItemsSource = AssignmentsCollection;

            // Template path — auto-fill z session memory
            TemplateBox.Text = BbsSessionState.LastTemplatePath ?? "";

            ContractNoBox.Text = initial.ContractNo    ?? "";
            Address1Box.Text   = initial.AddressLine1  ?? "";
            Address2Box.Text   = initial.AddressLine2  ?? "";
            Address3Box.Text   = initial.AddressLine3  ?? "";
            RevisionBox.Text   = initial.Revision      ?? "C1";
            PlotSuffixBox.Text = initial.PlotSuffix    ?? "";
        }

        private void OnTemplateBrowseClick(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title  = "Select BBS template (.xls or .xlsx)",
                Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
            };
            if (!string.IsNullOrWhiteSpace(TemplateBox.Text))
            {
                var dir = System.IO.Path.GetDirectoryName(TemplateBox.Text);
                if (System.IO.Directory.Exists(dir))
                    dlg.InitialDirectory = dir;
            }
            if (dlg.ShowDialog() == true)
                TemplateBox.Text = dlg.FileName;
        }

        private void OnOkClick(object sender, RoutedEventArgs e)
        {
            string template = TemplateBox.Text;
            if (string.IsNullOrWhiteSpace(template))
            {
                MessageBox.Show(
                    "Template BBS path is required.",
                    "Missing template",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }
            if (!System.IO.File.Exists(template))
            {
                MessageBox.Show(
                    "Template file not found:\n" + template,
                    "File not found",
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

            // Session memory — następnym razem auto-fill
            BbsSessionState.LastTemplatePath = template;
            SelectedTemplatePath = template;

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
