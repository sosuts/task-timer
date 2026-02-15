using System.Drawing;
using System.Windows;
using H.NotifyIcon;
using TaskTimer.Models;
using TaskTimer.Services;

namespace TaskTimer;

public partial class App : Application
{
    private TaskbarIcon? _notifyIcon;

    protected override void OnStartup(StartupEventArgs e)
    {
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
        Shutdown();
    }

    private static Icon CreateDefaultIcon()
    {
        // 16x16のシンプルなタイマーアイコンを動的に生成
        var bmp = new Bitmap(16, 16);
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
        return Icon.FromHandle(handle);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _notifyIcon?.Dispose();
        base.OnExit(e);
    }
}
