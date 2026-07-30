using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace UIRequirement;

/// <summary>
/// Interaction logic for AddMoreImages.xaml
/// </summary>
public partial class AddMoreImages : Window
{
    public ObservableCollection<string> AvailableImages { get; set; }
        = new();

    public ObservableCollection<string> ImagesAdded { get; set; }
        = new();

    private string? _selectedImagePath;

    public AddMoreImages()
    {
        InitializeComponent();

        cbAddMoreImages.ItemsSource = AvailableImages;
        lbImagesAdded.ItemsSource = ImagesAdded;

        LoadImages();
    }

    private void LoadImages()
    {
        // Sample data
        AvailableImages.Add(@"C:\Images\Image1.jpg");
        AvailableImages.Add(@"C:\Images\Image2.jpg");
        AvailableImages.Add(@"C:\Images\Image3.jpg");
    }

    private void cbAddMoreImages_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (cbAddMoreImages.SelectedItem is not string imagePath)
            return;

        _selectedImagePath = imagePath;

        try
        {
            pictureBox1.Source = new BitmapImage(
                new Uri(imagePath, UriKind.Absolute));
        }
        catch
        {
            pictureBox1.Source = null;
        }
    }

    private void btnAdd_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(tbImageName.Text))
        {
            MessageBox.Show("Please enter an image name.");
            return;
        }

        ImagesAdded.Add(tbImageName.Text);

        tbImageName.Clear();
    }

    private void btnClear_Click(object sender, RoutedEventArgs e)
    {
        ImagesAdded.Clear();

        tbImageName.Clear();

        pictureBox1.Source = null;

        cbAddMoreImages.SelectedIndex = -1;

        _selectedImagePath = null;
    }

    private void btnOK_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void btnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
