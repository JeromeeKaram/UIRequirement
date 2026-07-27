using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.ObjectModel;
using System.Windows;
using UIRequirement.Models;

namespace UIRequirement.ViewModels;

public partial class MainViewModel : ObservableObject
{
    [ObservableProperty]
    private string? findingTicketNo;

    [ObservableProperty]
    private string? esn;

    [ObservableProperty]
    private string? tsn;

    [ObservableProperty]
    private string? csn;

    [ObservableProperty]
    private string? damageDescription;

    public ObservableCollection<DamageInfo> Damages { get; } = new();

    [RelayCommand]
    private void Load()
    {
        MessageBox.Show("Load clicked");
    }

    [RelayCommand]
    private void Add()
    {
        MessageBox.Show("Add Clicked");
    }

    [RelayCommand]
    private void Clear()
    {
        MessageBox.Show("Clear clicked");
    }

    [RelayCommand]
    private void AddMoreImages()
    {
        MessageBox.Show("Add More Images");
    }

    [RelayCommand]
    private void Run()
    {
        MessageBox.Show("Run");
    }

    [RelayCommand]
    private void Close()
    {
        MessageBox.Show("Close clicked");
        //Application.Current.Shutdown();
    }
}
