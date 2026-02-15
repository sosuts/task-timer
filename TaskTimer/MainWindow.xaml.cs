using System.Windows;
using System.Windows.Input;

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
        Hide();
    }

    public void ForceClose()
    {
        Closing -= MainWindow_Closing;
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

        var needUpdate = false;

        // 右端がはみ出る場合
        if (Left + Width > virtualLeft + virtualWidth)
        {
            Left = virtualLeft + virtualWidth - Width;
            needUpdate = true;
        }

        // 下端がはみ出る場合
        if (Top + Height > virtualTop + virtualHeight)
        {
            Top = virtualTop + virtualHeight - Height;
            needUpdate = true;
        }

        // 左端がはみ出る場合
        if (Left < virtualLeft)
        {
            Left = virtualLeft;
            needUpdate = true;
        }

        // 上端がはみ出る場合
        if (Top < virtualTop)
        {
            Top = virtualTop;
            needUpdate = true;
        }
    }
}
