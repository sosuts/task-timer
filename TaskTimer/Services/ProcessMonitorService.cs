using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Automation;
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
    private string _lastDetectedBrowserUrl = string.Empty;
    private bool _disposed;

    // ブラウザプロセス名のリスト
    private static readonly HashSet<string> BrowserProcessNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "chrome", "msedge", "firefox", "brave", "opera", "iexplore"
    };

    /// <summary>
    /// タスクカテゴリが変更されたときに発火
    /// </summary>
    public event EventHandler<TaskDetectedEventArgs>? TaskDetected;

    /// <summary>
    /// 監視対象外のプロセスがアクティブになったときに発火
    /// </summary>
    public event EventHandler? TaskLost;

    /// <summary>
    /// 現在検知しているブラウザのURL
    /// </summary>
    public string LastDetectedBrowserUrl => _lastDetectedBrowserUrl;

    /// <summary>
    /// ブラウザURLが変化したときに発火
    /// </summary>
    public event EventHandler<string>? BrowserTitleChanged;

    public ProcessMonitorService(AppSettings settings)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        _timer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(Math.Max(1, settings.ProcessCheckIntervalSeconds))
        };
        _timer.Tick += CheckActiveProcess;
    }

    public void Start()
    {
        if (!_disposed)
            _timer.Start();
    }

    public void Stop() => _timer.Stop();

    private void CheckActiveProcess(object? sender, EventArgs e)
    {
        if (_disposed) return;

        try
        {
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return;

            GetWindowThreadProcessId(hwnd, out var pid);
            if (pid == 0) return;

            Process? process = null;
            try
            {
                process = Process.GetProcessById((int)pid);
            }
            catch (ArgumentException)
            {
                // プロセスが既に終了している
                return;
            }
            catch (InvalidOperationException)
            {
                // プロセス情報にアクセスできない
                return;
            }

            var processName = process.ProcessName;

            var sb = new System.Text.StringBuilder(512);
            GetWindowText(hwnd, sb, sb.Capacity);
            var windowTitle = sb.ToString();

            // ブラウザの場合、URLを取得してドメインマッピングをチェック
            var mapping = FindMapping(processName, windowTitle);

            if (mapping != null)
            {
                var category = mapping.Category;

                // ブラウザの場合、UIAutomationでURLを取得してドメインマッチを確認
                if (mapping.Category == TaskCategory.CodeReview && BrowserProcessNames.Contains(processName))
                {
                    var browserUrl = GetBrowserUrl(hwnd, processName);
                    
                    // URL情報を通知
                    if (_lastDetectedBrowserUrl != browserUrl)
                    {
                        _lastDetectedBrowserUrl = browserUrl;
                        BrowserTitleChanged?.Invoke(this, browserUrl);
                    }

                    var domainMapping = FindBrowserDomainMapping(browserUrl);
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
                    if (_currentDetectedCategory != category || _currentWindowTitle != browserUrl)
                    {
                        _currentDetectedCategory = category;
                        _currentWindowTitle = browserUrl;
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

                // ブラウザ以外の場合、URLをクリア
                if (_lastDetectedBrowserUrl != string.Empty)
                {
                    _lastDetectedBrowserUrl = string.Empty;
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

                // URLをクリア
                if (_lastDetectedBrowserUrl != string.Empty)
                {
                    _lastDetectedBrowserUrl = string.Empty;
                    BrowserTitleChanged?.Invoke(this, string.Empty);
                }
            }
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or UnauthorizedAccessException)
        {
            // Win32 APIやアクセス権限の問題は無視
        }
    }

    /// <summary>
    /// UIAutomationを使ってブラウザのURLを取得する
    /// </summary>
    private static string GetBrowserUrl(IntPtr hwnd, string processName)
    {
        try
        {
            var element = AutomationElement.FromHandle(hwnd);
            if (element == null) return string.Empty;

            // ブラウザごとに異なるアドレスバーの取得方法
            AutomationElement? urlBar = null;

            if (processName.Equals("chrome", StringComparison.OrdinalIgnoreCase) ||
                processName.Equals("msedge", StringComparison.OrdinalIgnoreCase) ||
                processName.Equals("brave", StringComparison.OrdinalIgnoreCase) ||
                processName.Equals("opera", StringComparison.OrdinalIgnoreCase))
            {
                // Chromium系ブラウザ: Edit コントロールを探す
                urlBar = element.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
            }
            else if (processName.Equals("firefox", StringComparison.OrdinalIgnoreCase))
            {
                // Firefox: ComboBox内のEditを探す
                var comboBox = element.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox));
                if (comboBox != null)
                {
                    urlBar = comboBox.FindFirst(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                }
                
                // 見つからない場合は直接Editを探す
                urlBar ??= element.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
            }

            if (urlBar != null)
            {
                // ValuePatternでURLを取得
                if (urlBar.TryGetCurrentPattern(ValuePattern.Pattern, out var pattern))
                {
                    var valuePattern = (ValuePattern)pattern;
                    return valuePattern.Current.Value;
                }
            }
        }
        catch (ElementNotAvailableException)
        {
            // ウィンドウが閉じられた等
        }
        catch (InvalidOperationException)
        {
            // UIAutomation要素にアクセスできない
        }

        return string.Empty;
    }

    /// <summary>
    /// BrowserDomainMappingsからURLにマッチするマッピングを探す
    /// </summary>
    private BrowserDomainMapping? FindBrowserDomainMapping(string url)
    {
        if (string.IsNullOrWhiteSpace(url)) return null;

        foreach (var dm in _settings.BrowserDomainMappings)
        {
            if (!string.IsNullOrWhiteSpace(dm.Domain) &&
                url.Contains(dm.Domain, StringComparison.OrdinalIgnoreCase))
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
        if (_disposed) return;
        _disposed = true;
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
