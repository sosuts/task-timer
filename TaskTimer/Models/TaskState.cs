namespace TaskTimer.Models;

/// <summary>
/// タスクの状態
/// </summary>
public enum TaskState
{
    /// <summary>実行中</summary>
    Running,

    /// <summary>一時停止中（アイドル検知など）</summary>
    Paused,

    /// <summary>停止済み</summary>
    Stopped
}
