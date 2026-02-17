using System.IO;
using System.Text;
using TaskTimer.Models;

namespace TaskTimer.Services;

/// <summary>
/// タスク記録をCSVファイルにエクスポートするサービス
/// </summary>
public static class CsvExportService
{
    public static string Export(IEnumerable<TaskRecord> records, string? outputDirectory = null)
    {
        var dir = GetSafeOutputDirectory(outputDirectory);
        Directory.CreateDirectory(dir);

        var fileName = $"TaskTimer_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
        var filePath = Path.Combine(dir, fileName);

        var sb = new StringBuilder();
        sb.AppendLine(TaskRecord.CsvHeader);

        foreach (var record in records)
        {
            sb.AppendLine(record.ToCsvLine());
        }

        File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        return filePath;
    }

    /// <summary>
    /// 出力ディレクトリを検証し、安全なパスを返す
    /// </summary>
    private static string GetSafeOutputDirectory(string? outputDirectory)
    {
        // デフォルトディレクトリ
        var defaultDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "TaskTimer");

        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            return defaultDir;
        }

        try
        {
            // パスを正規化してパストラバーサルを防止
            var fullPath = Path.GetFullPath(outputDirectory);

            // 危険な文字が含まれていないか確認
            var invalidChars = Path.GetInvalidPathChars();
            if (outputDirectory.IndexOfAny(invalidChars) >= 0)
            {
                return defaultDir;
            }

            // 相対パス（..）を含む場合は拒否
            if (outputDirectory.Contains(".."))
            {
                return defaultDir;
            }

            return fullPath;
        }
        catch
        {
            return defaultDir;
        }
    }
}
