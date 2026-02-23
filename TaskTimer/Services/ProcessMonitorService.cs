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
    private string _cachedVsCodeContext = string.Empty;  // VSCodeコンテキストのキャッシュ
    private DateTime _vsCodeContextLastFetched = DateTime.MinValue;
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
        {
            _timer.Start();
            CheckActiveProcess(null, EventArgs.Empty);
        }
    }

    /// <summary>
    /// 内部の検知状態をリセットする（次回チェック時に再検知が行われる）。
    /// このクラスは DispatcherTimer を使用するため、すべてのメソッドは UI スレッドから呼び出すこと。
    /// </summary>
    public void ResetDetectionState()
    {
        _currentDetectedCategory = null;
        _currentWindowTitle = string.Empty;
    }

    public void Stop() => _timer.Stop();

    private void CheckActiveProcess(object? sender, EventArgs e)
    {
        if (_disposed) return;

        try
        {
            // フォアグラウンドウィンドウ（ユーザーが操作中のウィンドウ）を取得
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero)
            {
                FireTaskLostIfNeeded();
                return;
            }

            GetWindowThreadProcessId(hwnd, out var pid);
            if (pid == 0)
            {
                FireTaskLostIfNeeded();
                return;
            }

            Process? process = null;
            try
            {
                process = Process.GetProcessById((int)pid);
            }
            catch (ArgumentException) { FireTaskLostIfNeeded(); return; }
            catch (InvalidOperationException) { FireTaskLostIfNeeded(); return; }

            var processName = process.ProcessName;
            var sb = new System.Text.StringBuilder(512);
            GetWindowText(hwnd, sb, sb.Capacity);
            var windowTitle = sb.ToString();

            var mapping = FindMapping(processName, windowTitle);
            if (mapping == null)
            {
                FireTaskLostIfNeeded();
                return;
            }

            var category = mapping.Category;

            // ブラウザの場合、UIAutomationでURLを取得してドメインマッチを確認
            if (mapping.Category == TaskCategory.CodeReview && BrowserProcessNames.Contains(processName))
            {
                var browserUrl = GetBrowserUrl(hwnd, processName);

                if (_lastDetectedBrowserUrl != browserUrl)
                {
                    _lastDetectedBrowserUrl = browserUrl;
                    BrowserTitleChanged?.Invoke(this, browserUrl);
                }

                var domainMapping = FindBrowserDomainMapping(browserUrl);
                if (domainMapping == null)
                {
                    FireTaskLostIfNeeded();
                    return;
                }

                if (_currentDetectedCategory != category || _currentWindowTitle != browserUrl)
                {
                    _currentDetectedCategory = category;
                    _currentWindowTitle = browserUrl;
                    var contextInfo = GetContextInfo(processName, windowTitle, browserUrl);
                    TaskDetected?.Invoke(this, new TaskDetectedEventArgs
                    {
                        Category = category,
                        WindowTitle = windowTitle,
                        ProcessName = processName,
                        DefaultLabel = domainMapping.TaskName,
                        ContextInfo = contextInfo,
                        ContextKey = contextInfo,
                        BrowserUrl = browserUrl
                    });
                }
                return;
            }

            // ブラウザ以外
            if (_lastDetectedBrowserUrl != string.Empty)
            {
                _lastDetectedBrowserUrl = string.Empty;
                BrowserTitleChanged?.Invoke(this, string.Empty);
            }

            if (_currentDetectedCategory != category || _currentWindowTitle != windowTitle)
            {
                _currentDetectedCategory = category;
                _currentWindowTitle = windowTitle;
                var contextInfo2 = GetContextInfo(processName, windowTitle, string.Empty);
                TaskDetected?.Invoke(this, new TaskDetectedEventArgs
                {
                    Category = category,
                    WindowTitle = windowTitle,
                    ProcessName = processName,
                    DefaultLabel = mapping.DefaultLabel,
                    ContextInfo = contextInfo2,
                    ContextKey = contextInfo2,
                    DocumentName = contextInfo2
                });
            }
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or UnauthorizedAccessException)
        {
            // Win32 APIやアクセス権限の問題は無視
        }
    }

    private void FireTaskLostIfNeeded()
    {
        if (_currentDetectedCategory != null)
        {
            _currentDetectedCategory = null;
            _currentWindowTitle = string.Empty;
            TaskLost?.Invoke(this, EventArgs.Empty);
        }

        if (_lastDetectedBrowserUrl != string.Empty)
        {
            _lastDetectedBrowserUrl = string.Empty;
            BrowserTitleChanged?.Invoke(this, string.Empty);
        }
    }

    /// <summary>
    /// ウィンドウタイトルからドキュメント名/ワークスペース名を抽出する
    /// </summary>
    private static string ExtractDocumentName(string processName, string windowTitle)
    {
        if (string.IsNullOrWhiteSpace(windowTitle))
            return string.Empty;

        // Visual Studio: "ファイル名 - プロジェクト名 - Microsoft Visual Studio"
        if (processName.Equals("devenv", StringComparison.OrdinalIgnoreCase))
        {
            var vsIdx = windowTitle.LastIndexOf(" - Microsoft Visual Studio", StringComparison.OrdinalIgnoreCase);
            if (vsIdx > 0)
                return windowTitle[..vsIdx].Trim();
            return windowTitle;
        }

        // VSCode: "ファイル名 - フォルダ名 - Visual Studio Code"
        if (processName.Equals("Code", StringComparison.OrdinalIgnoreCase))
        {
            var vscIdx = windowTitle.LastIndexOf(" - Visual Studio Code", StringComparison.OrdinalIgnoreCase);
            if (vscIdx > 0)
                return windowTitle[..vscIdx].Trim();
            return windowTitle;
        }

        // Word: "ドキュメント名 - Word" or "ドキュメント名 - Microsoft Word"
        if (processName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase))
        {
            var wordIdx = windowTitle.LastIndexOf(" - Word", StringComparison.OrdinalIgnoreCase);
            if (wordIdx < 0)
                wordIdx = windowTitle.LastIndexOf(" - Microsoft Word", StringComparison.OrdinalIgnoreCase);
            if (wordIdx > 0)
                return windowTitle[..wordIdx].Trim();
            return windowTitle;
        }

        // Excel: "ブック名 - Excel" or "ブック名 - Microsoft Excel"
        if (processName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase))
        {
            var excelIdx = windowTitle.LastIndexOf(" - Excel", StringComparison.OrdinalIgnoreCase);
            if (excelIdx < 0)
                excelIdx = windowTitle.LastIndexOf(" - Microsoft Excel", StringComparison.OrdinalIgnoreCase);
            if (excelIdx > 0)
                return windowTitle[..excelIdx].Trim();
            return windowTitle;
        }

        // TortoiseMerge: "ファイルパス - TortoiseMerge"
        if (processName.Equals("TortoiseMerge", StringComparison.OrdinalIgnoreCase))
        {
            var tmIdx = windowTitle.LastIndexOf(" - TortoiseMerge", StringComparison.OrdinalIgnoreCase);
            if (tmIdx > 0)
                return windowTitle[..tmIdx].Trim();
            return windowTitle;
        }

        return string.Empty;
    }

    /// <summary>
    /// タスクの同一性判定に使うコンテキストキーを抽出する。
    /// VSCode/VS: ワークスペース（フォルダ）名、Excel/Word/TortoiseMerge: ファイル名
    /// </summary>
    private static string ExtractContextKey(string processName, string windowTitle, string documentName)
    {
        if (string.IsNullOrWhiteSpace(documentName))
            return string.Empty;

        // VSCode: "ファイル名 - フォルダ名" → フォルダ名部分をコンテキストキーにする
        if (processName.Equals("Code", StringComparison.OrdinalIgnoreCase))
        {
            // documentName = "ファイル名 - フォルダ名" の場合、最後の " - " 以降がワークスペース名
            var lastSep = documentName.LastIndexOf(" - ", StringComparison.Ordinal);
            if (lastSep > 0)
                return documentName[(lastSep + 3)..].Trim();
            return documentName;
        }

        // Visual Studio: "ファイル名 - プロジェクト名" → プロジェクト名部分をコンテキストキーにする
        if (processName.Equals("devenv", StringComparison.OrdinalIgnoreCase))
        {
            var lastSep = documentName.LastIndexOf(" - ", StringComparison.Ordinal);
            if (lastSep > 0)
                return documentName[(lastSep + 3)..].Trim();
            return documentName;
        }

        // Excel/Word/TortoiseMerge: ファイル名そのものがコンテキストキー
        return documentName;
    }

    /// <summary>
    /// ブラウザURLからリポジトリパス（owner/repo）を抽出する。
    /// 例: https://github.com/user/repo/pull/123 → user/repo
    ///     https://gitlab.example.com/group/project/-/merge_requests/1 → group/project
    /// </summary>
    private static string ExtractBrowserRepoPath(string url)
    {
        if (string.IsNullOrWhiteSpace(url))
            return string.Empty;

        try
        {
            // スキーマがない場合は付与
            if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                url = "https://" + url;
            }

            var uri = new Uri(url);
            var segments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);

            // 最低2セグメント（owner/repo）が必要
            if (segments.Length >= 2)
            {
                return $"{segments[0]}/{segments[1]}";
            }

            // 1セグメントの場合はそのまま
            if (segments.Length == 1)
                return segments[0];
        }
        catch (UriFormatException)
        {
            // URL解析失敗
        }

        return url;
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

    /// <summary>
    /// プロセスに応じたコンテキスト情報を取得する
    /// </summary>
    private string GetContextInfo(string processName, string windowTitle, string browserUrl)
    {
        // ブラウザの場合はURLとタブ名を返す
        if (BrowserProcessNames.Contains(processName))
        {
            if (!string.IsNullOrEmpty(browserUrl))
            {
                // ウィンドウタイトルからタブ名を抽出（ブラウザは通常「タイトル - ブラウザ名」形式）
                var tabName = ExtractTabName(windowTitle, processName);
                return string.IsNullOrEmpty(tabName) ? browserUrl : $"{browserUrl} | {tabName}";
            }
            return windowTitle;
        }

        // Visual Studio
        if (processName.Equals("devenv", StringComparison.OrdinalIgnoreCase))
        {
            // ウィンドウタイトルから開いているプロジェクト/ソリューション名を抽出
            // 通常「ファイル名 - プロジェクト名 - Visual Studio」形式
            return ExtractFromWindowTitle(windowTitle, 1);
        }

        // VS Code
        if (processName.Equals("Code", StringComparison.OrdinalIgnoreCase))
        {
            // `code --status` の出力から開いているワークスペース名を取得
            // 出力形式: "ファイル名 - ワークスペース名 - ユーザー名 - Visual Studio Code"
            var vsCodeContext = GetVsCodeContext();
            if (!string.IsNullOrEmpty(vsCodeContext))
                return vsCodeContext;

            // フォールバック: ウィンドウタイトルから抽出
            var parts = SplitWindowTitle(windowTitle);
            if (parts.Length >= 2)
                return parts[^2];
            return windowTitle;
        }

        // Excel
        if (processName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase))
        {
            // ウィンドウタイトルからファイル名を抽出
            // 通常「ファイル名 - Excel」形式
            return ExtractFromWindowTitle(windowTitle, 0);
        }

        // Word
        if (processName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase))
        {
            // ウィンドウタイトルからファイル名を抽出
            // 通常「ファイル名 - Word」形式
            return ExtractFromWindowTitle(windowTitle, 0);
        }

        // TortoiseMerge
        if (processName.Equals("TortoiseMerge", StringComparison.OrdinalIgnoreCase))
        {
            // ウィンドウタイトルからファイル名を抽出
            // 通常「ファイル名 - TortoiseMerge」形式
            return ExtractFromWindowTitle(windowTitle, 0);
        }

        // その他のプロセスはウィンドウタイトルをそのまま返す
        return windowTitle;
    }

    /// <summary>
    /// ウィンドウタイトルを " - " で分割
    /// </summary>
    private static string[] SplitWindowTitle(string windowTitle)
    {
        return windowTitle.Split(new[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
    }

    /// <summary>
    /// `code --status` の出力からVS Codeで開いているワークスペース名を取得する。
    /// 結果は30秒間キャッシュし、UIスレッドのブロックを避けるため非同期で取得する。
    /// </summary>
    private string GetVsCodeContext()
    {
        // キャッシュが有効な場合はキャッシュを返す
        if ((DateTime.Now - _vsCodeContextLastFetched).TotalSeconds < 30)
            return _cachedVsCodeContext;

        // 非同期で更新（UIスレッドをブロックしない）
        _ = Task.Run(() =>
        {
            var result = FetchVsCodeContextFromCli();
            _cachedVsCodeContext = result;
            _vsCodeContextLastFetched = DateTime.Now;
        });

        // 初回はキャッシュが空のままフォールバックへ
        return _cachedVsCodeContext;
    }

    /// <summary>
    /// `code --status` を実際に実行してワークスペース名を取得する（バックグラウンド専用）
    /// </summary>
    private static string FetchVsCodeContextFromCli()
    {
        try
        {
            using var process = new System.Diagnostics.Process
            {
                StartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "code",
                    Arguments = "--status",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                }
            };
            process.Start();
            var output = process.StandardOutput.ReadToEnd();
            process.WaitForExit(3000);

            // "(... Visual Studio Code)" の括弧内を抽出
            var regex = new System.Text.RegularExpressions.Regex(@"\((.*?Visual Studio Code)\)");
            var matches = regex.Matches(output);

            // 重複排除してワークスペース名を収集
            var workspaces = matches
                .Cast<System.Text.RegularExpressions.Match>()
                .Select(m => m.Groups[1].Value)           // "ファイル名 - ワークスペース名 - ユーザー名 - Visual Studio Code"
                .Select(s =>
                {
                    var parts = s.Split(new[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
                    // インデックス1がワークスペース名（"ファイル名 - [ワークスペース名] - ..."）
                    return parts.Length >= 2 ? parts[1] : parts[0];
                })
                .Distinct()
                .ToList();

            if (workspaces.Count == 1)
                return workspaces[0];
            if (workspaces.Count > 1)
                return string.Join(", ", workspaces);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"code --status failed: {ex.Message}");
        }
        return string.Empty;
    }

    /// <summary>
    /// ウィンドウタイトルから指定インデックスの部分を抽出
    /// </summary>
    private static string ExtractFromWindowTitle(string windowTitle, int index)
    {
        var parts = SplitWindowTitle(windowTitle);
        if (parts.Length > index)
        {
            return parts[index];
        }
        return windowTitle;
    }

    /// <summary>
    /// ウィンドウタイトルからブラウザのタブ名を抽出
    /// </summary>
    private string ExtractTabName(string windowTitle, string processName)
    {
        if (string.IsNullOrEmpty(windowTitle)) return string.Empty;

        // ブラウザごとのタイトル形式に対応
        // Chrome/Edge/Brave: "タブ名 - Google Chrome" 形式
        // Firefox: "タブ名 - Mozilla Firefox" 形式
        var browserSuffixes = new[]
        {
            " - Google Chrome",
            " - Microsoft Edge",
            " - Brave",
            " - Mozilla Firefox",
            " - Opera"
        };

        foreach (var suffix in browserSuffixes)
        {
            if (windowTitle.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
            {
                return windowTitle[..^suffix.Length].Trim();
            }
        }

        return windowTitle;
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
    public string ContextInfo { get; set; } = string.Empty;
    public string ContextKey { get; set; } = string.Empty;
    public string BrowserUrl { get; set; } = string.Empty;
    public string DocumentName { get; set; } = string.Empty;
}
