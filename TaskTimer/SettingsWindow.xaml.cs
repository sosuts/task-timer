using System;
using System.Windows;
using TaskTimer.Models;
using TaskTimer.ViewModels;

namespace TaskTimer;

public partial class SettingsWindow : Window
{
    public SettingsWindow(AppSettings settings)
    {
        InitializeComponent();
        DataContext = new SettingsViewModel(settings);
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        if (Application.Current.MainWindow?.DataContext is MainViewModel vm)
        {
            vm.ReloadSettings();
        }
    }
}
