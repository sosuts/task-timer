using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using TaskTimer.Models;

namespace TaskTimer.Services;

/// <summary>
/// タスクセッションをJSONファイルに保存・読み込みするサービス
/// </summary>
public static class TaskSessionService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>
    /// 本日のセッションファイルパス
    /// </summary>
    private static string SessionFilePath =>
        Path.Combine(AppSettings.SettingsDir, $"session_{DateTime.Now:yyyyMMdd}.json");

    /// <summary>
    /// タスク一覧をJSONファイルに保存する
    /// </summary>
    public static void Save(IEnumerable<TaskRecord> tasks)
    {
        try
        {
            Directory.CreateDirectory(AppSettings.SettingsDir);
            var json = JsonSerializer.Serialize(tasks.ToList(), JsonOptions);
            File.WriteAllText(SessionFilePath, json);
        }
        catch
        {
            // 保存失敗は無視
        }
    }

    /// <summary>
    /// 本日のセッションファイルからタスク一覧を読み込む
    /// </summary>
    public static List<TaskRecord> Load()
    {
        var path = SessionFilePath;
        if (!File.Exists(path))
            return new List<TaskRecord>();

        try
        {
            var json = File.ReadAllText(path);
            var tasks = JsonSerializer.Deserialize<List<TaskRecord>>(json, JsonOptions) ?? new List<TaskRecord>();

            // 実行中・一時停止中のタスクは停止済みとして復元（アプリ終了で中断されたため）
            foreach (var task in tasks)
            {
                if (task.State == TaskState.Running || task.State == TaskState.Paused)
                {
                    task.State = TaskState.Stopped;
                    if (!task.EndTime.HasValue)
                        task.EndTime = task.StartTime + task.Elapsed;
                    task.PauseStartTime = null;
                }
            }

            return tasks;
        }
        catch
        {
            return new List<TaskRecord>();
        }
    }
}
