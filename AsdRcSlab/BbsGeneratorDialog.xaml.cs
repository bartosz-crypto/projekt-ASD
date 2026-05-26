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

        public BbsGenerationContext Result { get; private set; }

        public BbsGeneratorDialog(BbsGenerationContext initial)
        {
            InitializeComponent();
            DataContext = this;

            AssignmentsCollection =
                new ObservableCollection<BbsLayoutAssignment>(initial.Assignments);
            LayoutsGrid.ItemsSource = AssignmentsCollection;

            ContractNoBox.Text  = initial.ContractNo    ?? "";
            Address1Box.Text    = initial.AddressLine1  ?? "";
            Address2Box.Text    = initial.AddressLine2  ?? "";
            Address3Box.Text    = initial.AddressLine3  ?? "";
            RevisionBox.Text    = initial.Revision      ?? "C1";
            PlotSuffixBox.Text  = initial.PlotSuffix    ?? "";
        }

        private void OnOkClick(object sender, RoutedEventArgs e)
        {
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
