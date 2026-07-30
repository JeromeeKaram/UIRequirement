using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using UIRequirement;

namespace CTR_Form_Tool.Views
{
    public partial class UIHelp : Window
    {
        public UIHelp()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.Title = Utility.m_sToolName + " V" + Utility.m_sVersion;
            webBrowser1.Navigate(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin", "SOP_CTR Tool.pdf"));
        }

        private void PptMenuItem_Click(object sender, RoutedEventArgs e)
        {
            string pptFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin", "SOP_CTR Tool.pptx");

            Process.Start(new ProcessStartInfo
            {
                FileName = pptFile,
                UseShellExecute = true
            });
        }
    }
}