using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using H.NotifyIcon;
using TaskTimer.Models;
using TaskTimer.Services;

namespace TaskTimer;

public partial class App : Application
{
    private TaskbarIcon? _notifyIcon;
    private static Mutex? _mutex;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    protected override void OnStartup(StartupEventArgs e)
    {
        // 多重起動防止
        const string mutexName = "TaskTimer_SingleInstance_Mutex";
        _mutex = new Mutex(true, mutexName, out bool createdNew);

        if (!createdNew)
        {
            System.Windows.MessageBox.Show(
                LocalizationService.GetString("MessageAlreadyRunning"),
                LocalizationService.GetString("AppTitle"),
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            Shutdown();
            return;
        }

        base.OnStartup(e);

        var settings = AppSettings.Load();
        LocalizationService.ApplyLanguage(settings.Language);

        // システムトレイアイコンを作成
        _notifyIcon = new TaskbarIcon
        {
            ToolTipText = LocalizationService.GetString("AppTitle"),
        };

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
        if (MainWindow is MainWindow mw)
        {
            mw.ForceClose();
        }
        _mutex?.ReleaseMutex();
        _mutex?.Dispose();
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

    protected override void OnExit(ExitEventArgs e)
    {
        _notifyIcon?.Dispose();
        _mutex?.ReleaseMutex();
        _mutex?.Dispose();
        base.OnExit(e);
    }
}
