using System.Windows;
using System.Windows.Input;
using TaskTimer.Models;

namespace TaskTimer;

public partial class MainWindow : Window
{
    private const double MainWidth = 380;
    private double _mainPanelWidth = MainWidth;

    public MainWindow()
    {
        InitializeComponent();
        Closing += MainWindow_Closing;

        if (DataContext is ViewModels.MainViewModel vm)
        {
            vm.PropertyChanged += ViewModel_PropertyChanged;
            RestoreWindowState(vm.Settings);
        }
        DataContextChanged += (_, _) =>
        {
            if (DataContext is ViewModels.MainViewModel vm2)
            {
                vm2.PropertyChanged += ViewModel_PropertyChanged;
            }
        };

        Loaded += MainWindow_Loaded;
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        EnsureOnScreen();
    }

    private void RestoreWindowState(AppSettings settings)
    {
        if (!double.IsNaN(settings.WindowLeft) && !double.IsNaN(settings.WindowTop))
        {
            Left = settings.WindowLeft;
            Top = settings.WindowTop;
        }

        if (settings.WindowWidth > 0)
        {
            Width = settings.WindowWidth;
            _mainPanelWidth = settings.WindowWidth;
        }

        if (settings.WindowHeight > 0)
        {
            Height = settings.WindowHeight;
        }
    }

    private void SaveWindowState()
    {
        if (DataContext is ViewModels.MainViewModel vm)
        {
            var settings = vm.Settings;
            settings.WindowLeft = Left;
            settings.WindowTop = Top;
            settings.WindowWidth = vm.IsSidebarOpen ? _mainPanelWidth : Width;
            settings.WindowHeight = Height;
            settings.Save();
        }
    }

    private void ViewModel_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ViewModels.MainViewModel.IsSidebarOpen))
        {
            var vm = (ViewModels.MainViewModel)DataContext;
            if (vm.IsSidebarOpen)
            {
                _mainPanelWidth = ActualWidth;
                MaxWidth = double.PositiveInfinity;
                SidebarBorder.Visibility = Visibility.Visible;
                MainColumn.Width = new GridLength(1, GridUnitType.Star);
                SidebarColumn.Width = new GridLength(1, GridUnitType.Star);
                Width = _mainPanelWidth * 2;
            }
            else
            {
                SidebarBorder.Visibility = Visibility.Collapsed;
                SidebarColumn.Width = GridLength.Auto;
                MainColumn.Width = new GridLength(1, GridUnitType.Star);
                Width = _mainPanelWidth;
                MaxWidth = _mainPanelWidth + 40;
            }
            EnsureOnScreen();
        }
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        e.Cancel = true;
        SaveWindowState();
        Hide();
    }

    public void ForceClose()
    {
        Closing -= MainWindow_Closing;
        SaveWindowState();
        if (DataContext is ViewModels.MainViewModel vm)
        {
            vm.Dispose();
        }
        Close();
    }

    private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
            return;
        DragMove();
    }

    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState.Minimized;
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        SaveWindowState();
        Hide();
    }

    private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (DataContext is ViewModels.MainViewModel vm)
        {
            _mainPanelWidth = vm.IsSidebarOpen ? ActualWidth / 2 : ActualWidth;
        }
        EnsureOnScreen();
    }

    private void Window_LocationChanged(object? sender, EventArgs e)
    {
        // ドラッグ中の位置補正は行わない（ユーザーの意図を尊重）
    }

    /// <summary>
    /// ウィンドウが画面外にはみ出ないようにする（マルチディスプレイ対応）
    /// </summary>
    private void EnsureOnScreen()
    {
        // 仮想スクリーン全体の領域を取得
        var virtualLeft = SystemParameters.VirtualScreenLeft;
        var virtualTop = SystemParameters.VirtualScreenTop;
        var virtualWidth = SystemParameters.VirtualScreenWidth;
        var virtualHeight = SystemParameters.VirtualScreenHeight;

        // 右端がはみ出る場合
        if (Left + Width > virtualLeft + virtualWidth)
        {
            Left = virtualLeft + virtualWidth - Width;
        }

        // 下端がはみ出る場合
        if (Top + Height > virtualTop + virtualHeight)
        {
            Top = virtualTop + virtualHeight - Height;
        }

        // 左端がはみ出る場合
        if (Left < virtualLeft)
        {
            Left = virtualLeft;
        }

        // 上端がはみ出る場合
        if (Top < virtualTop)
        {
            Top = virtualTop;
        }
    }
}
