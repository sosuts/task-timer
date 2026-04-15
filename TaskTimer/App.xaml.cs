using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using H.NotifyIcon;
using Microsoft.Win32;
using TaskTimer.Models;
using TaskTimer.Services;

namespace TaskTimer;

public partial class App : Application
{
    private TaskbarIcon? _notifyIcon;
    private static Mutex? _mutex;
    private static bool _ownsMutex;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    protected override void OnStartup(StartupEventArgs e)
    {
        // 予期しない例外でプロセスが終了しないようにグローバルハンドラを登録する
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnDomainUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

        // スリープ復帰時にDispatcherTimerが正常に動作し続けるよう監視する
        SystemEvents.PowerModeChanged += OnPowerModeChanged;

        // 多重起動防止
        const string mutexName = "TaskTimer_SingleInstance_Mutex";
        _mutex = new Mutex(true, mutexName, out bool createdNew);

        if (!createdNew)
        {
            _mutex.Dispose();
            _mutex = null;
            System.Windows.MessageBox.Show(
                LocalizationService.GetString("MessageAlreadyRunning"),
                LocalizationService.GetString("AppTitle"),
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            Shutdown();
            return;
        }

        _ownsMutex = true;

        base.OnStartup(e);

        var settings = AppSettings.Load();
        LocalizationService.ApplyLanguage(settings.Language);

        // システムトレイアイコンを作成
        _notifyIcon = new TaskbarIcon
        {
            ToolTipText = LocalizationService.GetString("AppTitle"),
        };
        // コードで生成した場合はForceCreate()でシステムトレイへ登録する
        _notifyIcon.ForceCreate(enablesEfficiencyMode: false);

        // コンテキストメニュー
        var contextMenu = new System.Windows.Controls.ContextMenu();

        var showItem = new System.Windows.Controls.MenuItem { Header = LocalizationService.GetString("MenuShowWindow") };
        showItem.Click += (_, _) => ShowMainWindow();

        var exitItem = new System.Windows.Controls.MenuItem { Header = LocalizationService.GetString("MenuExit") };
        exitItem.Click += (_, _) => ExitApp();

        contextMenu.Items.Add(showItem);
        contextMenu.Items.Add(new System.Windows.Controls.Separator());
        contextMenu.Items.Add(exitItem);

        _notifyIcon.ContextMenu = contextMenu;
        _notifyIcon.TrayMouseDoubleClick += (_, _) => ShowMainWindow();

        // プログラムでアイコンを生成（外部ICOファイル不要）
        _notifyIcon.Icon = CreateDefaultIcon();
    }

    private void ShowMainWindow()
    {
        if (MainWindow == null)
        {
            MainWindow = new MainWindow();
        }
        MainWindow.Show();
        MainWindow.WindowState = WindowState.Normal;
        MainWindow.Activate();
    }

    private void ExitApp()
    {
        _notifyIcon?.Dispose();
        _notifyIcon = null;
        if (MainWindow is MainWindow mw)
        {
            mw.ForceClose();
        }
        Shutdown();
    }

    private static Icon CreateDefaultIcon()
    {
        // 16x16のシンプルなタイマーアイコンを動的に生成
        using var bmp = new Bitmap(16, 16);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        // 背景円（青）
        using var bgBrush = new SolidBrush(Color.FromArgb(26, 115, 232));
        g.FillEllipse(bgBrush, 1, 1, 14, 14);

        // 時計の針
        using var pen = new Pen(Color.White, 1.5f);
        g.DrawLine(pen, 8, 8, 8, 4);  // 分針
        g.DrawLine(pen, 8, 8, 11, 8); // 時針

        var handle = bmp.GetHicon();
        var icon = Icon.FromHandle(handle);
        // Icon.FromHandle は所有権を取得しないため、コピーを作成
        var clonedIcon = (Icon)icon.Clone();
        DestroyIcon(handle);
        return clonedIcon;
    }

    /// <summary>
    /// UIスレッドでの未処理例外を捕捉してプロセスを継続させる
    /// </summary>
    private static void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"[TaskTimer] DispatcherUnhandledException: {e.Exception}");
        e.Handled = true;
    }

    /// <summary>
    /// バックグラウンドスレッドでの未処理例外をログに記録する（プロセス終了は防げないが記録は残す）
    /// </summary>
    private static void OnDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"[TaskTimer] UnhandledException: {e.ExceptionObject}");
    }

    /// <summary>
    /// Task内の未観測例外を捕捉してプロセスを継続させる
    /// </summary>
    private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"[TaskTimer] UnobservedTaskException: {e.Exception}");
        e.SetObserved();
    }

    /// <summary>
    /// スリープ復帰時にDispatcherTimerが動作していることを確認する
    /// </summary>
    private void OnPowerModeChanged(object sender, PowerModeChangedEventArgs e)
    {
        if (e.Mode == PowerModes.Resume)
        {
            System.Diagnostics.Debug.WriteLine("[TaskTimer] System resumed from sleep.");
            // MainWindowのViewModelにスリープ復帰を通知してタイマーを再起動する
            Dispatcher.BeginInvoke(() =>
            {
                if (MainWindow?.DataContext is ViewModels.MainViewModel vm)
                {
                    vm.OnSystemResumed();
                }
            });
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        SystemEvents.PowerModeChanged -= OnPowerModeChanged;
        _notifyIcon?.Dispose();
        if (_ownsMutex)
        {
            _mutex?.ReleaseMutex();
            _mutex?.Dispose();
            _ownsMutex = false;
        }
        base.OnExit(e);
    }
}
