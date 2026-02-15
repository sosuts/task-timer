using System.Globalization;
using System.Linq;
using System.Windows;
using TaskTimer.Models;

namespace TaskTimer.Services;

public static class LocalizationService
{
    private const string DictionaryPrefix = "Resources/Strings";

    public static void ApplyLanguage(LanguagePreference language)
    {
        var culture = language == LanguagePreference.Japanese ? "ja-JP" : "en";
        var source = language == LanguagePreference.Japanese
            ? new Uri($"{DictionaryPrefix}.ja-JP.xaml", UriKind.Relative)
            : new Uri($"{DictionaryPrefix}.xaml", UriKind.Relative);

        var dictionaries = Application.Current.Resources.MergedDictionaries;
        var existing = dictionaries.FirstOrDefault(d => d.Source != null && d.Source.OriginalString.StartsWith(DictionaryPrefix, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            dictionaries.Remove(existing);
        }

        dictionaries.Add(new ResourceDictionary { Source = source });
        CultureInfo.CurrentUICulture = new CultureInfo(culture);
        CultureInfo.CurrentCulture = new CultureInfo(culture);
    }

    public static string GetString(string key)
    {
        return Application.Current.TryFindResource(key) as string ?? key;
    }
}
