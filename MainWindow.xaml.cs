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
        public MainWindow()
        {
            InitializeComponent();

            DataContext = new MainViewModel();
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
            MessageBox.Show(
                "Select the image that provides a close-up view of the damage area. "
                + "This image should clearly show the defect, crack, wear, dent, or other non-conformance details.",
                "Zoomed View",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void OverviewInfo_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "Select the image that shows the complete part or assembly. "
                + "This image should provide overall context and help identify the location of the damage relative to the entire component.",
                "Overview Image",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void PartInformationInfo_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "Select the image that contains part identification details such as the Part Number (P/N), Serial Number (S/N), nomenclature, ATA reference, or any markings required to uniquely identify the component.",
                "Part Information",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}