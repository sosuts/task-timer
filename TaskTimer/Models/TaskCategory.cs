namespace TaskTimer.Models;

/// <summary>
/// タスクのカテゴリ（どのアプリケーションでの作業か）
/// </summary>
public enum TaskCategory
{
    /// <summary>手動で追加されたタスク</summary>
    Manual,

    /// <summary>GitLabでのコードレビュー</summary>
    CodeReview,

    /// <summary>VSCode での実装作業</summary>
    VSCode,

    /// <summary>Visual Studio での実装作業</summary>
    VisualStudio,

    /// <summary>Word での仕様書作成</summary>
    Word,

    /// <summary>Excel での仕様書作成</summary>
    Excel,

    /// <summary>その他</summary>
    Other
}
