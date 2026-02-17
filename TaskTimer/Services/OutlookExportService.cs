using System.Runtime.InteropServices;
using TaskTimer.Models;

namespace TaskTimer.Services;

/// <summary>
/// タスク記録をOutlookの予定表にエクスポートするサービス
/// </summary>
public static class OutlookExportService
{
    // Outlook定数
    private const int OlFolderCalendar = 9;
    private const int OlAppointmentItem = 1;

    /// <summary>
    /// タスクをOutlookの指定予定表に登録する
    /// </summary>
    /// <param name="records">登録するタスク一覧</param>
    /// <param name="calendarName">対象の予定表名（空欄の場合は既定の予定表）</param>
    /// <returns>登録した件数</returns>
    public static int Export(IEnumerable<TaskRecord> records, string? calendarName = null)
    {
        dynamic? outlookApp = null;
        dynamic? ns = null;
        dynamic? calendarFolder = null;

        try
        {
            outlookApp = GetOutlookApplication();
            ns = outlookApp!.GetNamespace("MAPI");

            calendarFolder = FindOrCreateCalendar(ns, calendarName);

            int count = 0;
            foreach (var record in records)
            {
                if (record.StartTime == default) continue;

                dynamic appt = outlookApp.CreateItem(OlAppointmentItem);
                try
                {
                    appt.Subject = BuildSubject(record);
                    appt.Start = record.StartTime;
                    appt.End = record.EndTime ?? (record.StartTime + record.Elapsed);
                    appt.Body = BuildBody(record);
                    appt.ReminderSet = false;

                    // 指定予定表に移動
                    appt.Save();
                    if (calendarFolder != null)
                    {
                        dynamic moved = appt.Move(calendarFolder);
                        moved.Save();
                        ReleaseComObject(moved);
                    }

                    count++;
                }
                finally
                {
                    ReleaseComObject(appt);
                }
            }

            return count;
        }
        finally
        {
            ReleaseComObject(calendarFolder);
            ReleaseComObject(ns);
            // Outlookアプリケーション自体は解放しない（ユーザーが使用中の可能性）
        }
    }

    /// <summary>
    /// Outlookの利用可能な予定表名一覧を取得する
    /// </summary>
    public static List<string> GetCalendarNames()
    {
        var names = new List<string>();
        dynamic? outlookApp = null;
        dynamic? ns = null;
        dynamic? defaultCalendar = null;
        dynamic? folders = null;

        try
        {
            outlookApp = GetOutlookApplication();
            ns = outlookApp!.GetNamespace("MAPI");
            defaultCalendar = ns.GetDefaultFolder(OlFolderCalendar);

            // 既定の予定表
            names.Add((string)defaultCalendar.Name);

            // サブフォルダ（追加の予定表）
            folders = defaultCalendar.Folders;
            int folderCount = folders.Count;
            for (int i = 1; i <= folderCount; i++)
            {
                dynamic folder = folders[i];
                try
                {
                    names.Add((string)folder.Name);
                }
                finally
                {
                    ReleaseComObject(folder);
                }
            }
        }
        finally
        {
            ReleaseComObject(folders);
            ReleaseComObject(defaultCalendar);
            ReleaseComObject(ns);
        }

        return names;
    }

    private static dynamic GetOutlookApplication()
    {
        // まず実行中のOutlookを取得
        try
        {
            return GetActiveOutlookInstance();
        }
        catch
        {
            // 実行中でなければ新規起動
        }

        var outlookType = Type.GetTypeFromProgID("Outlook.Application");
        if (outlookType == null)
        {
            throw new InvalidOperationException(
                "Outlookがインストールされていません。\nOutlook is not installed.");
        }

        return Activator.CreateInstance(outlookType)
            ?? throw new InvalidOperationException(
                "Outlookの起動に失敗しました。\nFailed to start Outlook.");
    }

    /// <summary>
    /// 実行中のOutlookインスタンスを取得（.NET 8対応）
    /// </summary>
    private static dynamic GetActiveOutlookInstance()
    {
        var clsid = new Guid("0006F03A-0000-0000-C000-000000000046"); // Outlook.Application CLSID
        GetActiveObject(ref clsid, IntPtr.Zero, out var obj);
        return obj;
    }

    /// <summary>
    /// 指定名の予定表フォルダを探すか、なければ作成する
    /// </summary>
    private static dynamic? FindOrCreateCalendar(dynamic ns, string? calendarName)
    {
        dynamic defaultCalendar = ns.GetDefaultFolder(OlFolderCalendar);

        // 名前指定なし → 既定の予定表を使用
        if (string.IsNullOrWhiteSpace(calendarName))
        {
            return defaultCalendar;
        }

        // 既定の予定表名と一致する場合
        if (string.Equals((string)defaultCalendar.Name, calendarName, StringComparison.OrdinalIgnoreCase))
        {
            return defaultCalendar;
        }

        // サブフォルダから検索
        dynamic folders = defaultCalendar.Folders;
        try
        {
            int count = folders.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic folder = folders[i];
                if (string.Equals((string)folder.Name, calendarName, StringComparison.OrdinalIgnoreCase))
                {
                    ReleaseComObject(defaultCalendar);
                    return folder;
                }
                ReleaseComObject(folder);
            }
        }
        finally
        {
            ReleaseComObject(folders);
        }

        // 見つからなければ新しい予定表フォルダを作成
        dynamic newFolder = defaultCalendar.Folders.Add(calendarName, OlFolderCalendar);
        ReleaseComObject(defaultCalendar);
        return newFolder;
    }

    private static string BuildSubject(TaskRecord record)
    {
        var categoryTag = record.Category switch
        {
            TaskCategory.Manual => "📝",
            TaskCategory.CodeReview => "🔍",
            TaskCategory.VSCode => "💻",
            TaskCategory.VisualStudio => "🖥",
            TaskCategory.Word => "📄",
            TaskCategory.Excel => "📊",
            _ => "📁"
        };
        return $"{categoryTag} {record.TaskName}";
    }

    private static string BuildBody(TaskRecord record)
    {
        return $"タスク名: {record.TaskName}\n" +
               $"ラベル: {record.Label}\n" +
               $"カテゴリ: {record.Category}\n" +
               $"開始: {record.StartTime:yyyy-MM-dd HH:mm:ss}\n" +
               $"終了: {record.EndTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A"}\n" +
               $"経過時間: {record.Elapsed:hh\\:mm\\:ss}\n" +
               $"一時停止時間: {record.PausedDuration:hh\\:mm\\:ss}\n" +
               $"実質作業時間: {record.EffectiveElapsed:hh\\:mm\\:ss}";
    }

    private static void ReleaseComObject(object? obj)
    {
        if (obj != null)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch
            {
                // COMオブジェクトの解放失敗は無視
            }
        }
    }

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
}
