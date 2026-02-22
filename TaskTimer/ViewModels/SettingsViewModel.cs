using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TaskTimer.Models;
using TaskTimer.Services;

namespace TaskTimer.ViewModels;

public partial class SettingsViewModel : ObservableObject
{
    private readonly AppSettings _settings;

    [ObservableProperty]
    private string _gitLabDomain = string.Empty;

    [ObservableProperty]
    private int _idleThresholdMinutes;

    [ObservableProperty]
    private int _processCheckIntervalSeconds;

    [ObservableProperty]
    private bool _alwaysOnTop;

    [ObservableProperty]
    private string _csvOutputDirectory = string.Empty;

    [ObservableProperty]
    private string _outlookCalendarName = string.Empty;

    [ObservableProperty]
    private ObservableCollection<string> _outlookCalendarOptions = new();

    [ObservableProperty]
    private string _outlookCalendarStatus = string.Empty;

    [ObservableProperty]
    private string _saveStatusMessage = string.Empty;

    [ObservableProperty]
    private ObservableCollection<BrowserDomainMapping> _browserDomainMappings = new();

    public ObservableCollection<ProcessMapping> ProcessMappings { get; } = new();

    [ObservableProperty]
    private int _selectedFontSizeIndex;

    [ObservableProperty]
    private int _selectedLanguageIndex;

    public string[] FontSizeOptions =>
    [
        LocalizationService.GetString("FontSizeSmall"),
        LocalizationService.GetString("FontSizeMedium"),
        LocalizationService.GetString("FontSizeLarge")
    ];
    public string[] LanguageOptions =>
    [
        LocalizationService.GetString("LanguageEnglish"),
        LocalizationService.GetString("LanguageJapanese")
    ];

    public SettingsViewModel(AppSettings settings)
    {
        _settings = settings;
        GitLabDomain = settings.GitLabDomain;
        IdleThresholdMinutes = settings.IdleThresholdSeconds / 60;
        ProcessCheckIntervalSeconds = settings.ProcessCheckIntervalSeconds;
        AlwaysOnTop = settings.AlwaysOnTop;
        CsvOutputDirectory = settings.CsvOutputDirectory;
        OutlookCalendarName = settings.OutlookCalendarName;

        foreach (var m in settings.BrowserDomainMappings)
        {
            BrowserDomainMappings.Add(new BrowserDomainMapping { Domain = m.Domain, TaskName = m.TaskName });
        }

        foreach (var m in settings.ProcessMappings)
        {
            ProcessMappings.Add(new ProcessMapping
            {
                ProcessName = m.ProcessName,
                WindowTitleContains = m.WindowTitleContains,
                Category = m.Category,
                DefaultLabel = m.DefaultLabel
            });
        }

        SelectedFontSizeIndex = (int)settings.FontSize;
        SelectedLanguageIndex = (int)settings.Language;

        // Outlook予定表一覧を初期ロード
        LoadOutlookCalendars();
    }

    [RelayCommand]
    private void AddDomainMapping()
    {
        BrowserDomainMappings.Add(new BrowserDomainMapping { Domain = "", TaskName = "" });
    }

    [RelayCommand]
    private void RemoveDomainMapping(BrowserDomainMapping? mapping)
    {
        if (mapping != null)
        {
            BrowserDomainMappings.Remove(mapping);
        }
    }

    [RelayCommand]
    private void AddProcessMapping()
    {
        ProcessMappings.Add(new ProcessMapping { ProcessName = "", DefaultLabel = "", Category = TaskCategory.Other });
    }

    [RelayCommand]
    private void RemoveProcessMapping(ProcessMapping? mapping)
    {
        if (mapping != null)
        {
            ProcessMappings.Remove(mapping);
        }
    }

    [RelayCommand]
    private void RefreshOutlookCalendars()
    {
        LoadOutlookCalendars();
    }

    private void LoadOutlookCalendars()
    {
        OutlookCalendarStatus = "";
        try
        {
            var names = OutlookExportService.GetCalendarNames();
            OutlookCalendarOptions.Clear();
            foreach (var name in names)
            {
                OutlookCalendarOptions.Add(name);
            }

            if (OutlookCalendarOptions.Count > 0)
            {
                OutlookCalendarStatus = string.Format(
                    LocalizationService.GetString("OutlookCalendarFoundFormat"),
                    OutlookCalendarOptions.Count);
            }
        }
        catch (Exception ex)
        {
            OutlookCalendarOptions.Clear();
            OutlookCalendarStatus = LocalizationService.GetString("OutlookCalendarLoadFailed");
            System.Diagnostics.Debug.WriteLine($"Outlook calendar load failed: {ex.Message}");
        }
    }

    [RelayCommand]
    private void Save()
    {
        _settings.GitLabDomain = GitLabDomain;
        _settings.IdleThresholdSeconds = IdleThresholdMinutes * 60;
        _settings.ProcessCheckIntervalSeconds = ProcessCheckIntervalSeconds;
        _settings.AlwaysOnTop = AlwaysOnTop;
        _settings.CsvOutputDirectory = CsvOutputDirectory;
        _settings.OutlookCalendarName = OutlookCalendarName;
        _settings.FontSize = (FontSizePreference)SelectedFontSizeIndex;
        _settings.Language = (LanguagePreference)SelectedLanguageIndex;

        LocalizationService.ApplyLanguage(_settings.Language);

        _settings.BrowserDomainMappings.Clear();
        foreach (var m in BrowserDomainMappings)
        {
            if (!string.IsNullOrWhiteSpace(m.Domain))
            {
                _settings.BrowserDomainMappings.Add(new BrowserDomainMapping { Domain = m.Domain, TaskName = m.TaskName });
            }
        }

        _settings.ProcessMappings.Clear();
        foreach (var m in ProcessMappings)
        {
            if (!string.IsNullOrWhiteSpace(m.ProcessName))
            {
                _settings.ProcessMappings.Add(new ProcessMapping
                {
                    ProcessName = m.ProcessName,
                    WindowTitleContains = m.WindowTitleContains,
                    Category = m.Category,
                    DefaultLabel = m.DefaultLabel
                });
            }
        }

        _settings.Save();
        SaveStatusMessage = LocalizationService.GetString("SaveStatusMessage");
    }
}
