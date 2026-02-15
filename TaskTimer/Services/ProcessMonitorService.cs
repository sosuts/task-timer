using System.Diagnostics;
using System.Runtime.InteropServices;
using TaskTimer.Models;

namespace TaskTimer.Services;

/// <summary>
/// アクティブなプロセスを監視し、タスクの自動検知を行うサービス
/// </summary>
public class ProcessMonitorService : IDisposable
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);

    private readonly System.Windows.Threading.DispatcherTimer _timer;
    private readonly AppSettings _settings;
    private TaskCategory? _currentDetectedCategory;
    private string _currentWindowTitle = string.Empty;
    private string _lastDetectedBrowserTitle = string.Empty;

    /// <summary>
    /// タスクカテゴリが変更されたときに発火
    /// </summary>
    public event EventHandler<TaskDetectedEventArgs>? TaskDetected;

    /// <summary>
    /// 監視対象外のプロセスがアクティブになったときに発火
    /// </summary>
    public event EventHandler? TaskLost;

    /// <summary>
    /// 現在検知しているブラウザのウィンドウタイトル
    /// </summary>
    public string LastDetectedBrowserTitle => _lastDetectedBrowserTitle;

    /// <summary>
    /// ブラウザタイトルが変化したときに発火
    /// </summary>
    public event EventHandler<string>? BrowserTitleChanged;

    public ProcessMonitorService(AppSettings settings)
    {
        _settings = settings;
        _timer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(settings.ProcessCheckIntervalSeconds)
        };
        _timer.Tick += CheckActiveProcess;
    }

    public void Start() => _timer.Start();
    public void Stop() => _timer.Stop();

    private void CheckActiveProcess(object? sender, EventArgs e)
    {
        try
        {
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return;

            GetWindowThreadProcessId(hwnd, out var pid);
            var process = Process.GetProcessById((int)pid);
            var processName = process.ProcessName;

            var sb = new System.Text.StringBuilder(512);
            GetWindowText(hwnd, sb, sb.Capacity);
            var windowTitle = sb.ToString();

            // ブラウザの場合、ドメインマッピングをチェック
            var mapping = FindMapping(processName, windowTitle);

            if (mapping != null)
            {
                var category = mapping.Category;

                // ブラウザの場合、BrowserDomainMappingsでドメインマッチを確認
                if (mapping.Category == TaskCategory.CodeReview)
                {
                    // ブラウザタイトルを常に通知
                    if (_lastDetectedBrowserTitle != windowTitle)
                    {
                        _lastDetectedBrowserTitle = windowTitle;
                        BrowserTitleChanged?.Invoke(this, windowTitle);
                    }

                    var domainMapping = FindBrowserDomainMapping(windowTitle);
                    if (domainMapping == null)
                    {
                        // どのドメインにもマッチしないブラウザは無視
                        if (_currentDetectedCategory != null)
                        {
                            _currentDetectedCategory = null;
                            _currentWindowTitle = string.Empty;
                            TaskLost?.Invoke(this, EventArgs.Empty);
                        }
                        return;
                    }

                    // マッチしたドメインのタスク名を使用
                    if (_currentDetectedCategory != category || _currentWindowTitle != windowTitle)
                    {
                        _currentDetectedCategory = category;
                        _currentWindowTitle = windowTitle;
                        TaskDetected?.Invoke(this, new TaskDetectedEventArgs
                        {
                            Category = category,
                            WindowTitle = windowTitle,
                            ProcessName = processName,
                            DefaultLabel = domainMapping.TaskName
                        });
                    }
                    return;
                }

                if (_currentDetectedCategory != category || _currentWindowTitle != windowTitle)
                {
                    _currentDetectedCategory = category;
                    _currentWindowTitle = windowTitle;
                    TaskDetected?.Invoke(this, new TaskDetectedEventArgs
                    {
                        Category = category,
                        WindowTitle = windowTitle,
                        ProcessName = processName,
                        DefaultLabel = mapping.DefaultLabel
                    });
                }

                // ブラウザ以外の場合、タイトルをクリア
                if (_lastDetectedBrowserTitle != string.Empty)
                {
                    _lastDetectedBrowserTitle = string.Empty;
                    BrowserTitleChanged?.Invoke(this, string.Empty);
                }
            }
            else
            {
                if (_currentDetectedCategory != null)
                {
                    _currentDetectedCategory = null;
                    _currentWindowTitle = string.Empty;
                    TaskLost?.Invoke(this, EventArgs.Empty);
                }

                // ブラウザタイトルをクリア
                if (_lastDetectedBrowserTitle != string.Empty)
                {
                    _lastDetectedBrowserTitle = string.Empty;
                    BrowserTitleChanged?.Invoke(this, string.Empty);
                }
            }
        }
        catch
        {
            // プロセス情報取得失敗は無視
        }
    }

    /// <summary>
    /// BrowserDomainMappingsからウィンドウタイトルにマッチするマッピングを探す
    /// </summary>
    private BrowserDomainMapping? FindBrowserDomainMapping(string windowTitle)
    {
        foreach (var dm in _settings.BrowserDomainMappings)
        {
            if (!string.IsNullOrWhiteSpace(dm.Domain) &&
                windowTitle.Contains(dm.Domain, StringComparison.OrdinalIgnoreCase))
            {
                return dm;
            }
        }
        return null;
    }

    private ProcessMapping? FindMapping(string processName, string windowTitle)
    {
        foreach (var m in _settings.ProcessMappings)
        {
            if (string.Equals(m.ProcessName, processName, StringComparison.OrdinalIgnoreCase))
            {
                if (m.WindowTitleContains == null ||
                    windowTitle.Contains(m.WindowTitleContains, StringComparison.OrdinalIgnoreCase))
                {
                    return m;
                }
            }
        }
        return null;
    }

    public void Dispose()
    {
        _timer.Stop();
        GC.SuppressFinalize(this);
    }
}

public class TaskDetectedEventArgs : EventArgs
{
    public TaskCategory Category { get; set; }
    public string WindowTitle { get; set; } = string.Empty;
    public string ProcessName { get; set; } = string.Empty;
    public string DefaultLabel { get; set; } = string.Empty;
}
