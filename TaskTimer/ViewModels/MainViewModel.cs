using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TaskTimer.Models;
using TaskTimer.Services;

namespace TaskTimer.ViewModels;

public partial class MainViewModel : ObservableObject, IDisposable
{
    private AppSettings _settings;
    private IdleDetectionService _idleService;
    private ProcessMonitorService _processMonitor;
    private readonly DispatcherTimer _tickTimer;
    private readonly DispatcherTimer _clockTimer;
    private int _autoSaveCounter;

    [ObservableProperty]
    private ObservableCollection<TaskRecord> _tasks = new();

    [ObservableProperty]
    private TaskRecord? _activeTask;

    [ObservableProperty]
    private string _newTaskName = string.Empty;

    [ObservableProperty]
    private string _newTaskLabel = string.Empty;

    [ObservableProperty]
    private TaskCategory _newTaskCategory = TaskCategory.Manual;

    [ObservableProperty]
    private bool _isAlwaysOnTop;

    [ObservableProperty]
    private bool _isAutoDetectEnabled = true;

    [ObservableProperty]
    private bool _isIdleDetectEnabled = true;

    [ObservableProperty]
    private string _statusMessage = LocalizationService.GetString("StatusReady");

    [ObservableProperty]
    private string _currentElapsedDisplay = "00:00";

    [ObservableProperty]
    private string _autoDetectStatus = "";

    [ObservableProperty]
    private int _sessionCount;

    [ObservableProperty]
    private string _totalFocusTimeDisplay = "0m";

    [ObservableProperty]
    private string _currentTimeDisplay = "00:00:00";

    [ObservableProperty]
    private string _currentDateDisplay = "";

    [ObservableProperty]
    private bool _isSidebarOpen;

    [ObservableProperty]
    private string _detectedBrowserTitle = "";

    [ObservableProperty]
    private double _fontSizeSmall = 11;

    [ObservableProperty]
    private double _fontSizeMedium = 14;

    [ObservableProperty]
    private double _fontSizeLarge = 18;

    [ObservableProperty]
    private double _fontSizeClock = 56;

    [ObservableProperty]
    private double _fontSizeDate = 15;

    [ObservableProperty]
    private double _fontSizeElapsed = 18;

    public Array CategoryValues => Enum.GetValues(typeof(TaskCategory));

    public AppSettings Settings => _settings;

    public MainViewModel()
    {
        _settings = AppSettings.Load();
        _isAlwaysOnTop = _settings.AlwaysOnTop;
        LocalizationService.ApplyLanguage(_settings.Language);
        ApplyFontSize(_settings.FontSize);
        StatusMessage = LocalizationService.GetString("StatusReady");
        _idleService = new IdleDetectionService(_settings.IdleThresholdSeconds);
        _idleService.IdleStarted += OnIdleStarted;
        _idleService.IdleEnded += OnIdleEnded;

        _processMonitor = new ProcessMonitorService(_settings);
        _processMonitor.TaskDetected += OnTaskDetected;
        _processMonitor.TaskDetectionCycleCompleted += OnTaskDetectionCycleCompleted;
        _processMonitor.TaskLost += OnTaskLost;
        _processMonitor.BrowserTitleChanged += OnBrowserTitleChanged;

        // 毎秒タイマー更新
        _tickTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        _tickTimer.Tick += OnTick;
        _tickTimer.Start();

        // 時計用タイマー
        _clockTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
        _clockTimer.Tick += OnClockTick;
        _clockTimer.Start();

        _idleService.Start();
        _processMonitor.Start();

        // 前回セッションのタスクを復元
        var savedTasks = TaskSessionService.Load();
        foreach (var t in savedTasks)
            Tasks.Add(t);

        UpdateClock();
    }

    /// <summary>
    /// 設定変更後にサービスを再構成する
    /// </summary>
    public void ReloadSettings()
    {
        _settings = AppSettings.Load();
        IsAlwaysOnTop = _settings.AlwaysOnTop;
        LocalizationService.ApplyLanguage(_settings.Language);
        ApplyFontSize(_settings.FontSize);

        _idleService.Dispose();
        _idleService = new IdleDetectionService(_settings.IdleThresholdSeconds);
        _idleService.IdleStarted += OnIdleStarted;
        _idleService.IdleEnded += OnIdleEnded;
        _idleService.Start();

        _processMonitor.Dispose();
        _processMonitor = new ProcessMonitorService(_settings);
        _processMonitor.TaskDetected += OnTaskDetected;
        _processMonitor.TaskDetectionCycleCompleted += OnTaskDetectionCycleCompleted;
        _processMonitor.TaskLost += OnTaskLost;
        _processMonitor.BrowserTitleChanged += OnBrowserTitleChanged;
        _processMonitor.Start();
    }

    private void OnClockTick(object? sender, EventArgs e)
    {
        UpdateClock();
    }

    private void UpdateClock()
    {
        var now = DateTime.Now;
        CurrentTimeDisplay = now.ToString("HH:mm:ss");
        CurrentDateDisplay = now.ToString("yyyy/MM/dd (ddd)");
    }

    private void OnTick(object? sender, EventArgs e)
    {
        foreach (var runningTask in Tasks.Where(t => t.State == TaskState.Running))
        {
            runningTask.Elapsed = DateTime.Now - runningTask.StartTime - runningTask.PausedDuration;
        }

        if (ActiveTask is { State: TaskState.Running })
        {
            var ts = ActiveTask.Elapsed;
            CurrentElapsedDisplay = ts.TotalHours >= 1
                ? ts.ToString(@"h\:mm\:ss")
                : ts.ToString(@"mm\:ss");
        }

        UpdateTotalFocusTime();

        // 1分ごとにセッションを自動保存
        _autoSaveCounter++;
        if (_autoSaveCounter >= 60)
        {
            _autoSaveCounter = 0;
            TaskSessionService.Save(Tasks);
        }
    }

    private void UpdateTotalFocusTime()
    {
        var total = TimeSpan.Zero;
        foreach (var t in Tasks)
        {
            total += t.Elapsed;
        }

        if (total.TotalHours >= 1)
            TotalFocusTimeDisplay = $"{(int)total.TotalHours}h {total.Minutes:D2}m";
        else
            TotalFocusTimeDisplay = $"{(int)total.TotalMinutes}m";

        SessionCount = Tasks.Count;
    }

    private void OnIdleStarted(object? sender, EventArgs e)
    {
        if (!IsIdleDetectEnabled) return;
        if (ActiveTask is { State: TaskState.Running })
        {
            ActiveTask.State = TaskState.Paused;
            ActiveTask.PauseStartTime = DateTime.Now;
            StatusMessage = LocalizationService.GetString("StatusIdlePaused");
            AutoDetectStatus = LocalizationService.GetString("StatusIdle");
        }
    }

    private void OnIdleEnded(object? sender, EventArgs e)
    {
        if (!IsIdleDetectEnabled) return;
        if (ActiveTask is { State: TaskState.Paused, PauseStartTime: not null })
        {
            ActiveTask.PausedDuration += DateTime.Now - ActiveTask.PauseStartTime.Value;
            ActiveTask.PauseStartTime = null;
            ActiveTask.State = TaskState.Running;
            StatusMessage = $"{ActiveTask.TaskName}";
            AutoDetectStatus = $"{ActiveTask.Label}";
        }
    }

    private void OnTaskDetected(object? sender, TaskDetectedEventArgs e)
    {
        if (!IsAutoDetectEnabled) return;

        var existingRunning = Tasks.FirstOrDefault(t =>
            t.State == TaskState.Running &&
            t.Category == e.Category &&
            string.Equals(t.ContextKey, e.ContextKey, StringComparison.OrdinalIgnoreCase));

        if (existingRunning != null)
        {
            AutoDetectStatus = e.DefaultLabel;
            existingRunning.ProcessName = e.ProcessName;
            existingRunning.DetectedUrl = e.BrowserUrl;
            existingRunning.DetectedTabTitle = e.WindowTitle;
            existingRunning.DetectedDocumentName = e.DocumentName;
            if (string.IsNullOrEmpty(existingRunning.ContextKey))
                existingRunning.ContextKey = e.ContextKey;
            ActiveTask = existingRunning;
            return;
        }

        var existingPaused = Tasks.FirstOrDefault(t =>
            t.State == TaskState.Paused &&
            t.Category == e.Category &&
            string.Equals(t.ContextKey, e.ContextKey, StringComparison.OrdinalIgnoreCase));

        if (existingPaused != null)
        {
            if (existingPaused.PauseStartTime != null)
            {
                existingPaused.PausedDuration += DateTime.Now - existingPaused.PauseStartTime.Value;
                existingPaused.PauseStartTime = null;
            }

            existingPaused.State = TaskState.Running;
            existingPaused.ProcessName = e.ProcessName;
            existingPaused.DetectedUrl = e.BrowserUrl;
            existingPaused.DetectedTabTitle = e.WindowTitle;
            existingPaused.DetectedDocumentName = e.DocumentName;
            if (string.IsNullOrEmpty(existingPaused.ContextKey))
                existingPaused.ContextKey = e.ContextKey;

            ActiveTask = existingPaused;
            StatusMessage = $"{existingPaused.TaskName}";
            AutoDetectStatus = e.DefaultLabel;
            return;
        }

        // 同じカテゴリ＆同じコンテキストのタスクが実行中なら詳細情報のみ更新
        if (ActiveTask is { State: TaskState.Running } &&
            ActiveTask.Category == e.Category &&
            string.Equals(ActiveTask.ContextKey, e.ContextKey, StringComparison.OrdinalIgnoreCase))
        {
            AutoDetectStatus = e.DefaultLabel;
            ActiveTask.ProcessName = e.ProcessName;
            ActiveTask.DetectedUrl = e.BrowserUrl;
            ActiveTask.DetectedTabTitle = e.WindowTitle;
            ActiveTask.DetectedDocumentName = e.DocumentName;
            if (string.IsNullOrEmpty(ActiveTask.ContextKey))
                ActiveTask.ContextKey = e.ContextKey;
            return;
        }

        // 一時停止中で同一コンテキストなら再開して再利用
        if (ActiveTask is { State: TaskState.Paused } &&
            ActiveTask.Category == e.Category &&
            string.Equals(ActiveTask.ContextKey, e.ContextKey, StringComparison.OrdinalIgnoreCase))
        {
            if (ActiveTask.PauseStartTime != null)
            {
                ActiveTask.PausedDuration += DateTime.Now - ActiveTask.PauseStartTime.Value;
                ActiveTask.PauseStartTime = null;
            }

            ActiveTask.State = TaskState.Running;
            ActiveTask.ProcessName = e.ProcessName;
            ActiveTask.DetectedUrl = e.BrowserUrl;
            ActiveTask.DetectedTabTitle = e.WindowTitle;
            ActiveTask.DetectedDocumentName = e.DocumentName;
            if (string.IsNullOrEmpty(ActiveTask.ContextKey))
                ActiveTask.ContextKey = e.ContextKey;

            StatusMessage = $"{ActiveTask.TaskName}";
            AutoDetectStatus = e.DefaultLabel;
            return;
        }

        // 同じカテゴリ＆同じコンテキストで直近の停止済みタスクがあれば再開
        var recentStopped = FindRecentStoppedTask(e.Category, e.ContextKey);
        if (recentStopped != null)
        {
            // 停止済みタスクを再開
            recentStopped.State = TaskState.Running;
            recentStopped.EndTime = null;
            recentStopped.PauseStartTime = null;
            recentStopped.ProcessName = e.ProcessName;
            recentStopped.ContextInfo = e.ContextInfo;
            recentStopped.ContextKey = string.IsNullOrEmpty(recentStopped.ContextKey)
                ? e.ContextKey
                : recentStopped.ContextKey;
            recentStopped.DetectedUrl = e.BrowserUrl;
            recentStopped.DetectedTabTitle = e.WindowTitle;
            recentStopped.DetectedDocumentName = e.DocumentName;
            ActiveTask = recentStopped;
            StatusMessage = $"{recentStopped.TaskName}";
            AutoDetectStatus = e.DefaultLabel;
            return;
        }

        // コンテキストキーに基づいたタスク名を生成
        var contextLabel = !string.IsNullOrEmpty(e.ContextKey)
            ? $"{e.DefaultLabel} ({e.ContextKey})"
            : e.DefaultLabel;

        var task = new TaskRecord
        {
            TaskName = $"[Auto] {contextLabel}",
            Label = e.DefaultLabel,
            Category = e.Category,
            State = TaskState.Running,
            StartTime = DateTime.Now,
            ProcessName = e.ProcessName,
            ContextInfo = e.ContextInfo,
            ContextKey = e.ContextKey,
            DetectedUrl = e.BrowserUrl,
            DetectedTabTitle = e.WindowTitle,
            DetectedDocumentName = e.DocumentName
        };

        Tasks.Add(task);
        ActiveTask = task;
        StatusMessage = $"{task.TaskName}";
        AutoDetectStatus = e.DefaultLabel;
        CurrentElapsedDisplay = "00:00";
    }

    /// <summary>
    /// 同じカテゴリ＆同じコンテキストキーで最近停止されたタスクを探す（5分以内）
    /// </summary>
    private TaskRecord? FindRecentStoppedTask(TaskCategory category, string contextKey)
    {
        var threshold = TimeSpan.FromMinutes(5);
        TaskRecord? best = null;

        foreach (var t in Tasks)
        {
            if (t.State == TaskState.Stopped &&
                t.Category == category &&
                string.Equals(t.ContextKey, contextKey, StringComparison.OrdinalIgnoreCase) &&
                t.EndTime.HasValue &&
                (DateTime.Now - t.EndTime.Value) < threshold)
            {
                if (best == null || t.EndTime > best.EndTime)
                    best = t;
            }
        }

        return best;
    }

    private void OnTaskLost(object? sender, EventArgs e)
    {
        if (!IsAutoDetectEnabled) return;

        AutoDetectStatus = string.Empty;
    }

    private void OnTaskDetectionCycleCompleted(object? sender, DetectedTaskKeysEventArgs e)
    {
        if (!IsAutoDetectEnabled) return;

        var now = DateTime.Now;
        var toStop = Tasks
            .Where(t => t.State == TaskState.Running)
            .Where(t => t.TaskName.StartsWith("[Auto]", StringComparison.Ordinal))
            .Where(t => !e.Keys.Contains(BuildDetectionKey(t.Category, t.ContextKey)))
            .ToList();

        foreach (var task in toStop)
        {
            StopTaskRecord(task, now);
        }

        if (ActiveTask != null && toStop.Contains(ActiveTask))
        {
            ActiveTask = Tasks
                .Where(t => t.State == TaskState.Running)
                .OrderByDescending(t => t.StartTime)
                .FirstOrDefault();

            if (ActiveTask == null)
            {
                StatusMessage = LocalizationService.GetString("StatusReady");
                CurrentElapsedDisplay = "00:00";
            }
        }

        if (toStop.Count > 0)
        {
            TaskSessionService.Save(Tasks);
        }
    }

    private static string BuildDetectionKey(TaskCategory category, string contextKey)
    {
        return $"{category}|{contextKey}";
    }

    private void OnBrowserTitleChanged(object? sender, string title)
    {
        DetectedBrowserTitle = title;
    }

    [RelayCommand]
    private void StartOrPause()
    {
        if (ActiveTask is { State: TaskState.Running })
        {
            PauseResumeTask();
        }
        else if (ActiveTask is { State: TaskState.Paused })
        {
            PauseResumeTask();
        }
        else
        {
            if (string.IsNullOrWhiteSpace(NewTaskName))
            {
                NewTaskName = $"Task {Tasks.Count + 1}";
            }
            StartTask();
        }
    }

    [RelayCommand]
    private void StartTask()
    {
        if (string.IsNullOrWhiteSpace(NewTaskName)) return;

        if (ActiveTask is { State: TaskState.Running or TaskState.Paused })
        {
            StopCurrentTask();
        }

        var task = new TaskRecord
        {
            TaskName = NewTaskName,
            Label = NewTaskLabel,
            Category = NewTaskCategory,
            State = TaskState.Running,
            StartTime = DateTime.Now
        };

        Tasks.Add(task);
        ActiveTask = task;
        StatusMessage = $"{task.TaskName}";
        CurrentElapsedDisplay = "00:00";

        NewTaskName = string.Empty;
        NewTaskLabel = string.Empty;
        TaskSessionService.Save(Tasks);
    }

    [RelayCommand]
    private void StopTask()
    {
        if (ActiveTask is { State: TaskState.Running or TaskState.Paused })
        {
            StopCurrentTask();
            StatusMessage = LocalizationService.GetString("StatusReady");
            CurrentElapsedDisplay = "00:00";
        }
    }

    [RelayCommand]
    private void PauseResumeTask()
    {
        if (ActiveTask == null) return;

        if (ActiveTask.State == TaskState.Running)
        {
            ActiveTask.State = TaskState.Paused;
            ActiveTask.PauseStartTime = DateTime.Now;
            StatusMessage = string.Format(LocalizationService.GetString("StatusPausedFormat"), ActiveTask.TaskName);
        }
        else if (ActiveTask.State == TaskState.Paused && ActiveTask.PauseStartTime != null)
        {
            ActiveTask.PausedDuration += DateTime.Now - ActiveTask.PauseStartTime.Value;
            ActiveTask.PauseStartTime = null;
            ActiveTask.State = TaskState.Running;
            StatusMessage = $"{ActiveTask.TaskName}";
        }
    }

    [RelayCommand]
    private void DeleteTask(TaskRecord? task)
    {
        if (task == null) return;

        // 確認ダイアログ
        if (_settings.ConfirmOnDelete)
        {
            var result = MessageBox.Show(
                string.Format(LocalizationService.GetString("ConfirmDeleteTaskFormat"), task.TaskName),
                LocalizationService.GetString("ConfirmDeleteTitle"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return;
        }

        if (task == ActiveTask)
        {
            StopCurrentTask();
            ActiveTask = null;
            StatusMessage = LocalizationService.GetString("StatusReady");
            CurrentElapsedDisplay = "00:00";
        }
        Tasks.Remove(task);
        TaskSessionService.Save(Tasks);
    }

    [RelayCommand]
    private void ExportCsv()
    {
        try
        {
            var path = CsvExportService.Export(Tasks, _settings.CsvOutputDirectory);
            StatusMessage = LocalizationService.GetString("StatusCsvExported");
            MessageBox.Show(string.Format(LocalizationService.GetString("MessageCsvExportedFormat"), path),
                LocalizationService.GetString("MessageCsvExportedTitle"), MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(string.Format(LocalizationService.GetString("MessageCsvExportFailedFormat"), ex.Message),
                LocalizationService.GetString("MessageErrorTitle"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void ExportToOutlook()
    {
        if (!Tasks.Any())
        {
            MessageBox.Show(LocalizationService.GetString("MessageOutlookNoTasks"),
                LocalizationService.GetString("MessageOutlookExportedTitle"), MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            var calendarName = string.IsNullOrWhiteSpace(_settings.OutlookCalendarName)
                ? null
                : _settings.OutlookCalendarName;
            var count = OutlookExportService.Export(Tasks, calendarName);
            var displayName = calendarName ?? "(既定)";
            StatusMessage = LocalizationService.GetString("StatusOutlookExported");
            MessageBox.Show(
                string.Format(LocalizationService.GetString("MessageOutlookExportedFormat"), count, displayName),
                LocalizationService.GetString("MessageOutlookExportedTitle"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                string.Format(LocalizationService.GetString("MessageOutlookExportFailedFormat"), ex.Message),
                LocalizationService.GetString("MessageErrorTitle"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void OpenSettings()
    {
        var settingsWindow = new SettingsWindow(_settings);
        settingsWindow.Owner = Application.Current.MainWindow;
        settingsWindow.ShowDialog();
        ReloadSettings();
    }

    [RelayCommand]
    private void ToggleSidebar()
    {
        IsSidebarOpen = !IsSidebarOpen;
    }

    [RelayCommand]
    private void ClearCompletedTasks()
    {
        var stopped = Tasks.Where(t => t.State == TaskState.Stopped).ToList();
        if (stopped.Count == 0) return;

        // 確認ダイアログ
        if (_settings.ConfirmOnDelete)
        {
            var result = MessageBox.Show(
                string.Format(LocalizationService.GetString("ConfirmClearCompletedFormat"), stopped.Count),
                LocalizationService.GetString("ConfirmDeleteTitle"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return;
        }

        foreach (var t in stopped)
        {
            Tasks.Remove(t);
        }
        TaskSessionService.Save(Tasks);
    }

    private void StopCurrentTask()
    {
        if (ActiveTask == null) return;

        StopTaskRecord(ActiveTask, DateTime.Now);
        ActiveTask = null;
        TaskSessionService.Save(Tasks);
    }

    private static void StopTaskRecord(TaskRecord task, DateTime endTime)
    {
        if (task.State != TaskState.Running && task.State != TaskState.Paused)
            return;

        if (task.State == TaskState.Paused && task.PauseStartTime != null)
        {
            task.PausedDuration += endTime - task.PauseStartTime.Value;
            task.PauseStartTime = null;
        }

        task.State = TaskState.Stopped;
        task.EndTime = endTime;
        task.Elapsed = task.EndTime.Value - task.StartTime - task.PausedDuration;
    }

    private void ApplyFontSize(FontSizePreference pref)
    {
        double baseSize = pref switch
        {
            FontSizePreference.Small => 16,
            FontSizePreference.Large => 24,
            _ => 20
        };
        FontSizeSmall = baseSize * 0.78;
        FontSizeMedium = baseSize;
        FontSizeLarge = baseSize * 1.28;
        FontSizeClock = baseSize * 4.0;
        FontSizeDate = baseSize * 1.14;
        FontSizeElapsed = baseSize * 1.28;
    }

    partial void OnIsAlwaysOnTopChanged(bool value)
    {
        if (Application.Current.MainWindow != null)
        {
            Application.Current.MainWindow.Topmost = value;
        }
    }

    partial void OnIsAutoDetectEnabledChanged(bool value)
    {
        if (value)
        {
            // 再有効化時に内部状態をリセットして現在のプロセスを即時再検知する
            _processMonitor.ResetDetectionState();
        }
    }

    public void Dispose()
    {
        _tickTimer.Stop();
        _clockTimer.Stop();
        TaskSessionService.Save(Tasks);
        _idleService.Dispose();
        _processMonitor.Dispose();
        GC.SuppressFinalize(this);
    }
}
