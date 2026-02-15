using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace TaskTimer.Models;

/// <summary>
/// アプリケーション設定
/// </summary>
public class AppSettings
{
    /// <summary>アイドル判定までの秒数</summary>
    public int IdleThresholdSeconds { get; set; } = 300;

    /// <summary>プロセス監視間隔（秒）</summary>
    public int ProcessCheckIntervalSeconds { get; set; } = 5;

    /// <summary>常に最前面に表示</summary>
    public bool AlwaysOnTop { get; set; } = false;

    /// <summary>言語設定</summary>
    public LanguagePreference Language { get; set; } =
        CultureInfo.CurrentUICulture.TwoLetterISOLanguageName.Equals("ja", StringComparison.OrdinalIgnoreCase)
            ? LanguagePreference.Japanese
            : LanguagePreference.English;

    /// <summary>GitLabオンプレミスのドメイン（後方互換用）</summary>
    public string GitLabDomain { get; set; } = "gitlab.example.com";

    /// <summary>ブラウザドメイン→タスク名のマッピング一覧</summary>
    public List<BrowserDomainMapping> BrowserDomainMappings { get; set; } = new()
    {
        new() { Domain = "gitlab.example.com", TaskName = "コードレビュー" },
        new() { Domain = "github.com", TaskName = "GitHub" }
    };

    /// <summary>フォントサイズ設定</summary>
    public FontSizePreference FontSize { get; set; } = FontSizePreference.Medium;

    /// <summary>CSV出力先ディレクトリ</summary>
    public string CsvOutputDirectory { get; set; } = "";

    /// <summary>プロセスマッピング一覧</summary>
    public List<ProcessMapping> ProcessMappings { get; set; } = new()
    {
        new() { ProcessName = "chrome", WindowTitleContains = null, Category = TaskCategory.CodeReview, DefaultLabel = "コードレビュー" },
        new() { ProcessName = "msedge", WindowTitleContains = null, Category = TaskCategory.CodeReview, DefaultLabel = "コードレビュー" },
        new() { ProcessName = "firefox", WindowTitleContains = null, Category = TaskCategory.CodeReview, DefaultLabel = "コードレビュー" },
        new() { ProcessName = "brave", WindowTitleContains = null, Category = TaskCategory.CodeReview, DefaultLabel = "コードレビュー" },
        new() { ProcessName = "Code", WindowTitleContains = null, Category = TaskCategory.VSCode, DefaultLabel = "VSCode作業" },
        new() { ProcessName = "devenv", WindowTitleContains = null, Category = TaskCategory.VisualStudio, DefaultLabel = "Visual Studio作業" },
        new() { ProcessName = "WINWORD", WindowTitleContains = null, Category = TaskCategory.Word, DefaultLabel = "Word作業" },
        new() { ProcessName = "EXCEL", WindowTitleContains = null, Category = TaskCategory.Excel, DefaultLabel = "Excel作業" },
    };

    internal static readonly string SettingsDir = Path.GetDirectoryName(
        Environment.ProcessPath ?? AppContext.BaseDirectory)!;

    private static readonly string SettingsPath = Path.Combine(SettingsDir, "settings.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    public void Save()
    {
        Directory.CreateDirectory(SettingsDir);
        var json = JsonSerializer.Serialize(this, JsonOptions);
        File.WriteAllText(SettingsPath, json);
    }

    public static AppSettings Load()
    {
        if (!File.Exists(SettingsPath))
        {
            var defaults = new AppSettings();
            defaults.Save();
            return defaults;
        }

        try
        {
            var json = File.ReadAllText(SettingsPath);
            var settings = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
            settings.MigrateDefaults();
            return settings;
        }
        catch
        {
            return new AppSettings();
        }
    }

    /// <summary>
    /// 既知のデフォルトプロセスマッピングが不足していれば追加する
    /// </summary>
    private void MigrateDefaults()
    {
        var defaults = new AppSettings();
        var changed = false;

        foreach (var def in defaults.ProcessMappings)
        {
            var exists = ProcessMappings.Any(m =>
                string.Equals(m.ProcessName, def.ProcessName, StringComparison.OrdinalIgnoreCase));
            if (!exists)
            {
                ProcessMappings.Add(def);
                changed = true;
            }
        }

        if (changed)
        {
            Save();
        }
    }
}
