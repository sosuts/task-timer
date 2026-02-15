using System.Globalization;
using System.Windows;
using System.Windows.Data;
using TaskTimer.Models;
using TaskTimer.Services;

namespace TaskTimer.Converters;

/// <summary>
/// TaskState → 表示文字列
/// </summary>
public class TaskStateToStringConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is TaskState state ? state switch
        {
            TaskState.Running => LocalizationService.GetString("TaskStateRunning"),
            TaskState.Paused => LocalizationService.GetString("TaskStatePaused"),
            TaskState.Stopped => LocalizationService.GetString("TaskStateStopped"),
            _ => value.ToString() ?? ""
        } : "";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// TaskState → 背景色
/// </summary>
public class TaskStateToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is TaskState state ? state switch
        {
            TaskState.Running => "#2196F3",
            TaskState.Paused => "#FF9800",
            TaskState.Stopped => "#9E9E9E",
            _ => "#9E9E9E"
        } : "#9E9E9E";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// TaskCategory → 表示文字列
/// </summary>
public class TaskCategoryToStringConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is TaskCategory cat ? cat switch
        {
            TaskCategory.Manual => LocalizationService.GetString("TaskCategoryManual"),
            TaskCategory.CodeReview => LocalizationService.GetString("TaskCategoryCodeReview"),
            TaskCategory.VSCode => LocalizationService.GetString("TaskCategoryVSCode"),
            TaskCategory.VisualStudio => LocalizationService.GetString("TaskCategoryVisualStudio"),
            TaskCategory.Word => LocalizationService.GetString("TaskCategoryWord"),
            TaskCategory.Excel => LocalizationService.GetString("TaskCategoryExcel"),
            TaskCategory.Other => LocalizationService.GetString("TaskCategoryOther"),
            _ => value.ToString() ?? ""
        } : "";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// TaskCategory → アイコン色
/// </summary>
public class TaskCategoryToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is TaskCategory cat ? cat switch
        {
            TaskCategory.Manual => "#607D8B",
            TaskCategory.CodeReview => "#E91E63",
            TaskCategory.VSCode => "#2196F3",
            TaskCategory.VisualStudio => "#9C27B0",
            TaskCategory.Word => "#1565C0",
            TaskCategory.Excel => "#2E7D32",
            TaskCategory.Other => "#795548",
            _ => "#607D8B"
        } : "#607D8B";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// Bool → Visibility
/// </summary>
public class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is true ? Visibility.Visible : Visibility.Collapsed;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => value is Visibility.Visible;
}

/// <summary>
/// Null → Visibility (nullでないなら表示)
/// </summary>
public class NotNullToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value != null ? Visibility.Visible : Visibility.Collapsed;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// TimeSpan → 表示文字列
/// </summary>
public class TimeSpanToStringConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is TimeSpan ts ? ts.ToString(@"hh\:mm\:ss") : "00:00:00";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// DateTime → 表示文字列
/// </summary>
public class DateTimeToStringConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value switch
        {
            DateTime dt => dt.ToString("HH:mm:ss"),
            _ => ""
        };
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// Nullable DateTime → 表示文字列
/// </summary>
public class NullableDateTimeToStringConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is DateTime dt ? dt.ToString("HH:mm:ss") : "-";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}
