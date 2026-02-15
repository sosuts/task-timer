namespace TaskTimer.Models;

/// <summary>
/// ブラウザのドメインとタスク名のマッピング
/// </summary>
public class BrowserDomainMapping
{
    /// <summary>監視対象のドメイン（ブラウザのタイトルに含まれるか判定）</summary>
    public string Domain { get; set; } = string.Empty;

    /// <summary>検知時に付けるタスク名</summary>
    public string TaskName { get; set; } = string.Empty;
}
