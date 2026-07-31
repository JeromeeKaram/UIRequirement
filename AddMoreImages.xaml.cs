using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
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
    public ObservableCollection<KeyValuePair<string, string>> ExistingImages { get; set; }
        = new();

    public ObservableCollection<KeyValuePair<string, string>> ImagesAdded { get; set; }
        = new();

    private KeyValuePair<string, string>? selectedExistingImage;
    public KeyValuePair<string, string>? SelectedExistingImage
    {
        get => selectedExistingImage;
        set => selectedExistingImage = value;
    }

    public AddMoreImages()
    {
        InitializeComponent();
        DataContext = this;
    }

    private void cmbAddMoreImages_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (cmbAddMoreImages.SelectedItem is not KeyValuePair<string, string> image)
            return;

        selectedExistingImage = image;

        try
        {
            pictureBox1.Source = new BitmapImage(
                new Uri(image.Value, UriKind.Absolute));

            txtImageName.Text = image.Key.Split(".")[0];
        }
        catch
        {
            pictureBox1.Source = null;
        }
    }

    private void btnAdd_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtImageName.Text))
        {
            MessageBox.Show("Please enter an image name.");
            return;
        }

        ImagesAdded.Add(new KeyValuePair<string, string>(txtImageName.Text, selectedExistingImage?.Value ?? string.Empty));

        txtImageName.Clear();
    }

    private void btnClear_Click(object sender, RoutedEventArgs e)
    {
        ImagesAdded.Clear();
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

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        ExistingImages.Clear();
        var findingFolder = AppDomain.CurrentDomain.BaseDirectory + "\\bin\\images" + "\\Finding\\";

        string[] imageFiles = Directory.GetFiles(findingFolder, "*.*")
                                                .Where(file => file.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)
                                                || file.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase)
                                                || file.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
                                                || file.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase)
                                                || file.EndsWith(".gif", StringComparison.OrdinalIgnoreCase)).ToArray();

        //Show in the Drop down
        if (imageFiles != null && imageFiles.Length > 0)
        {
            foreach (string img in imageFiles)
            {
                ExistingImages.Add(new KeyValuePair<string, string>(System.IO.Path.GetFileName(img), img));
            }
        }
    }
}
