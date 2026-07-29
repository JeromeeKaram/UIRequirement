using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using UIRequirement.ViewModels;

namespace UIRequirement
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainViewModel _vm { get; }
        public MainWindow()
        {
            InitializeComponent();

            _vm = new MainViewModel();
            DataContext = _vm;
        }

        private void MenuExit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void MenuAbout_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "CTR Form Automation Tool V3.1\n\nDeveloped using WPF and Material Design.",
                "About",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void ZoomedViewInfo_Click(object sender, RoutedEventArgs e)
        {
            ShowImage(_vm.SelectedZoomedView);
        }

        private void OverviewInfo_Click(object sender, RoutedEventArgs e)
        {
            ShowImage(_vm.SelectedOverviewImage);
        }

        private void PartInformationInfo_Click(object sender, RoutedEventArgs e)
        {
            ShowImage(_vm.SelectedPartInformation);
        }

        private void ShowImage(string? imagePath)
        {
            if (string.IsNullOrWhiteSpace(imagePath))
            {
                MessageBox.Show("Please select an image first.");
                return;
            }

            if (!System.IO.File.Exists(imagePath))
            {
                MessageBox.Show("The selected image could not be found.");
                return;
            }

            var window = new ImageWindow(imagePath);
            window.Owner = this;
            window.ShowDialog();
        }

        //private void OverviewImages_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    if (sender is ComboBox combo &&
        //        combo.SelectedValue is string imagePath)
        //    {
        //        //ShowImage(imagePath);
        //        _vm.SelectedOverviewImage = combo.SelectedValue.ToString();

        //        // Reset the selection
        //        //combo.SelectedIndex = -1;
        //        // or combo.SelectedItem = null;
        //    }
        //}

        //private void ZoomedViews_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    if (sender is ComboBox combo &&
        //        combo.SelectedValue is string imagePath)
        //    {

        //        _vm.SelectedZoomedView = combo.SelectedValue.ToString();
        //        //ShowImage(imagePath);

        //        // Reset the selection
        //        //combo.SelectedIndex = -1;
        //        // or combo.SelectedItem = null;
        //    }
        //}

        //private void PartsInformation_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    if (sender is ComboBox combo &&
        //        combo.SelectedValue is string imagePath)
        //    {
        //        //ShowImage(imagePath);
        //        _vm.SelectedPartInformation = combo.SelectedValue.ToString();
        //        // Reset the selection
        //        //combo.SelectedIndex = -1;
        //        // or combo.SelectedItem = null;
        //    }
        //}
    }
}