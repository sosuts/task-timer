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
        var dir = string.IsNullOrWhiteSpace(outputDirectory)
            ? Path.Combine(AppSettings.SettingsDir, "TaskTimer")
            : outputDirectory;

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
}
