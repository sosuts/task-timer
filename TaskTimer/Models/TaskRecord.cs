using CommunityToolkit.Mvvm.ComponentModel;

namespace TaskTimer.Models;

/// <summary>
/// 1つのタスク計測レコード
/// </summary>
public partial class TaskRecord : ObservableObject
{
    [ObservableProperty]
    private string _id = Guid.NewGuid().ToString("N")[..8];

    [ObservableProperty]
    private string _taskName = string.Empty;

    [ObservableProperty]
    private string _label = string.Empty;

    [ObservableProperty]
    private TaskCategory _category = TaskCategory.Manual;

    [ObservableProperty]
    private TaskState _state = TaskState.Stopped;

    [ObservableProperty]
    private DateTime _startTime;

    [ObservableProperty]
    private DateTime? _endTime;

    [ObservableProperty]
    private TimeSpan _elapsed = TimeSpan.Zero;

    [ObservableProperty]
    private TimeSpan _pausedDuration = TimeSpan.Zero;

    /// <summary>
    /// 一時停止開始時刻（一時停止中のみ有効）
    /// </summary>
    public DateTime? PauseStartTime { get; set; }

    /// <summary>
    /// 検知元のプロセス名（例: chrome, devenv, Code）
    /// </summary>
    [ObservableProperty]
    private string _processName = string.Empty;

    /// <summary>
    /// ブラウザの場合: 検知時のURL
    /// </summary>
    [ObservableProperty]
    private string _detectedUrl = string.Empty;

    /// <summary>
    /// ブラウザの場合: 検知時のタブタイトル（ウィンドウタイトル）
    /// </summary>
    [ObservableProperty]
    private string _detectedTabTitle = string.Empty;

    /// <summary>
    /// VS/VSCode/Office/TortoiseMerge: 開いているファイル名またはワークスペース名
    /// </summary>
    [ObservableProperty]
    private string _detectedDocumentName = string.Empty;

    /// <summary>
    /// タスクの同一性を判定するためのコンテキストキー。
    /// VSCode/VS: ワークスペース名、Excel/Word: ファイル名、ブラウザ: リポジトリパス
    /// </summary>
    [ObservableProperty]
    private string _contextKey = string.Empty;

    /// <summary>
    /// 実質作業時間を算出
    /// </summary>
    public TimeSpan EffectiveElapsed => Elapsed - PausedDuration;

    /// <summary>
    /// CSV出力用の文字列
    /// </summary>
    public string ToCsvLine()
    {
        return $"\"{ Id}\",\"{Escape(TaskName)}\",\"{Escape(Label)}\",\"{Category}\"," +
               $"\"{ State}\",\"{StartTime:yyyy-MM-dd HH:mm:ss}\"," +
               $"\"{ EndTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? ""}\",\"{Elapsed:hh\\:mm\\:ss}\"," +
               $"\"{ PausedDuration:hh\\:mm\\:ss}\",\"{EffectiveElapsed:hh\\:mm\\:ss}\"," +
               $"\"{ Escape(ProcessName)}\",\"{Escape(DetectedUrl)}\",\"{Escape(DetectedTabTitle)}\",\"{Escape(DetectedDocumentName)}\"";
    }

    public static string CsvHeader =>
        "\"ID\",\"タスク名\",\"ラベル\",\"カテゴリ\",\"状態\",\"開始時刻\",\"終了時刻\",\"経過時間\",\"一時停止時間\",\"実質作業時間\",\"プロセス名\",\"URL\",\"タブ名\",\"ドキュメント名\"";

    private static string Escape(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        return value.Replace("\"", "\"\"");
    }
}
