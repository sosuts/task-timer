namespace TaskTimer.Models;

/// <summary>
/// プロセス監視の設定（どのプロセスをどのカテゴリに紐づけるか）
/// </summary>
public class ProcessMapping
{
    public string ProcessName { get; set; } = string.Empty;
    public string? WindowTitleContains { get; set; }
    public TaskCategory Category { get; set; }
    public string DefaultLabel { get; set; } = string.Empty;
}
