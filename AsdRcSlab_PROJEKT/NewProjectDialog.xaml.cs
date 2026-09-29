using System;
using System.Windows;
using System.Windows.Controls;

namespace AsdRcSlab
{
    public partial class NewProjectDialog : Window
    {
        public ProjectData Result { get; private set; }

        private static readonly string[] HystoolCodes =
        {
            "DK90","DK165","DK165","DK225","DK225","DK300","DK300",
            "DK375","DK375","DK450","DK450","DK525","DK525","DK600"
        };

        public NewProjectDialog(ProjectData existing = null)
        {
            InitializeComponent();
            TxtDate.Text = DateTime.Today.ToString("yyyy-MM-dd");

            if (existing != null)
            {
                TxtProjectName.Text = existing.ProjectName;
                TxtClientName.Text  = existing.ClientName;
                TxtDRWNumber.Text   = existing.DRWNumber;
                TxtRevision.Text    = existing.Revision;
                TxtDate.Text        = existing.ProjectDate;

                // Ustaw ComboBox na zapisana grubosc
                foreach (ComboBoxItem item in CmbSlab.Items)
                {
                    if (item.Content.ToString() == existing.SlabThickness.ToString())
                    {
                        CmbSlab.SelectedItem = item;
                        break;
                    }
                }
            }
        }

        private void CmbSlab_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TxtHystool == null) return;
            int idx = CmbSlab.SelectedIndex;
            TxtHystool.Text = (idx >= 0 && idx < HystoolCodes.Length)
                ? HystoolCodes[idx]
                : "—";
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            // Walidacja
            if (string.IsNullOrWhiteSpace(TxtProjectName.Text))
            {
                ShowError("Nazwa projektu jest wymagana.");
                TxtProjectName.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(TxtDRWNumber.Text))
            {
                ShowError("DRW Number jest wymagany.");
                TxtDRWNumber.Focus();
                return;
            }

            var selected = (CmbSlab.SelectedItem as ComboBoxItem)?.Content?.ToString();
            if (!int.TryParse(selected, out int thickness)) thickness = 300;

            Result = new ProjectData
            {
                ProjectName    = TxtProjectName.Text.Trim(),
                ClientName     = TxtClientName.Text.Trim(),
                DRWNumber      = TxtDRWNumber.Text.Trim(),
                Revision       = string.IsNullOrWhiteSpace(TxtRevision.Text) ? "P01" : TxtRevision.Text.Trim(),
                SlabThickness  = thickness,
                FCK            = 28,
                KSpring        = 212333,
                ProjectDate    = TxtDate.Text.Trim()
            };

            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void ShowError(string msg)
        {
            TxtError.Text       = msg;
            TxtError.Visibility = Visibility.Visible;
        }
    }
}
