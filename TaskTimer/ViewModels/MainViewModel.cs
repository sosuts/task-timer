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

        _idleService = new IdleDetectionService(_settings.IdleThresholdSeconds);
        _idleService.IdleStarted += OnIdleStarted;
        _idleService.IdleEnded += OnIdleEnded;

        _processMonitor = new ProcessMonitorService(_settings);
        _processMonitor.TaskDetected += OnTaskDetected;
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
        if (ActiveTask is { State: TaskState.Running })
        {
            ActiveTask.Elapsed = DateTime.Now - ActiveTask.StartTime - ActiveTask.PausedDuration;
            var ts = ActiveTask.Elapsed;
            CurrentElapsedDisplay = ts.TotalHours >= 1
                ? ts.ToString(@"h\:mm\:ss")
                : ts.ToString(@"mm\:ss");
        }

        UpdateTotalFocusTime();
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
            return;
        }

        // 同じカテゴリ＆同じコンテキストで直近の停止済みタスクがあれば再開
        var recentStopped = FindRecentStoppedTask(e.Category, e.ContextKey);
        if (recentStopped != null)
        {
            // 現在のタスクを停止
            if (ActiveTask is { State: TaskState.Running or TaskState.Paused })
            {
                StopCurrentTask();
            }

            // 停止済みタスクを再開
            recentStopped.State = TaskState.Running;
            recentStopped.EndTime = null;
            recentStopped.PauseStartTime = null;
            recentStopped.ProcessName = e.ProcessName;
            recentStopped.ContextInfo = e.ContextInfo;
            ActiveTask = recentStopped;
            StatusMessage = $"{recentStopped.TaskName}";
            AutoDetectStatus = e.DefaultLabel;
            return;
        }

        if (ActiveTask is { State: TaskState.Running or TaskState.Paused })
        {
            StopCurrentTask();
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
            ContextInfo = e.ContextInfo
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

        AutoDetectStatus = "";

        if (ActiveTask is { State: TaskState.Running } &&
            ActiveTask.TaskName.StartsWith("[Auto]"))
        {
            StopCurrentTask();
            StatusMessage = LocalizationService.GetString("StatusReady");
            CurrentElapsedDisplay = "00:00";
        }
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
    }

    private void StopCurrentTask()
    {
        if (ActiveTask == null) return;

        if (ActiveTask.State == TaskState.Paused && ActiveTask.PauseStartTime != null)
        {
            ActiveTask.PausedDuration += DateTime.Now - ActiveTask.PauseStartTime.Value;
            ActiveTask.PauseStartTime = null;
        }

        ActiveTask.State = TaskState.Stopped;
        ActiveTask.EndTime = DateTime.Now;
        ActiveTask.Elapsed = ActiveTask.EndTime.Value - ActiveTask.StartTime - ActiveTask.PausedDuration;
        ActiveTask = null;
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

    public void Dispose()
    {
        _tickTimer.Stop();
        _clockTimer.Stop();
        _idleService.Dispose();
        _processMonitor.Dispose();
        GC.SuppressFinalize(this);
    }
}
