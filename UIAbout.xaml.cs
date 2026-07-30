using System.Reflection;
using System.Windows;
using UIRequirement;

namespace CTR_Form_Tool.Views
{
    public partial class UIAbout : Window
    {
        public UIAbout()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                lbVersion.Text = "Version : " + Utility.m_sVersion;
                lbDate.Text = "Last Release : " + Utility.m_sReleaseDate;
                this.Title  = Utility.m_sToolName + " V" + Utility.m_sVersion;
                lblPurpose.Text = "The tool’s primary purpose is to auto generation of CTR (Customer Technical Report) by using the redmine damage information";
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }

        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}