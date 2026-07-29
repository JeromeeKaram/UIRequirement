using System;
using System.Windows;
using System.Windows.Media.Imaging;

namespace UIRequirement;

public partial class ImageWindow : Window
{
    public ImageWindow(string imagePath)
    {
        InitializeComponent();

        imgPreview.Source = new BitmapImage(new Uri(imagePath));
    }
}