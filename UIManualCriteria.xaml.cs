using System.Collections.Generic;
using System.Windows;

namespace UIRequirement.Views;

public partial class UIManualCriteria : Window
{
    public string ManualCriteria_Selected { get; private set; } = string.Empty;

    public UIManualCriteria(List<string> criteriaList)
    {
        InitializeComponent();

        foreach (string item in criteriaList)
        {
            cbMC.Items.Add(item);
        }

        if (cbMC.Items.Count > 0)
        {
            cbMC.SelectedIndex = 0;
        }

        Loaded += UIManualCriteria_Load;
    }

    private void UIManualCriteria_Load(object sender, RoutedEventArgs e)
    {
        // Equivalent of WinForms Load event
    }

    private void btnOK_Click(object sender, RoutedEventArgs e)
    {
        if (cbMC.SelectedItem != null)
        {
            ManualCriteria_Selected = cbMC.SelectedItem.ToString() ?? string.Empty;
        }

        DialogResult = true;
        Close();
    }

    private void btnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}