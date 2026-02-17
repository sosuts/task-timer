namespace TaskTimer.Models;

/// <summary>
/// Browser domain to task name mapping
/// </summary>
public class BrowserDomainMapping
{
    /// <summary>Domain to monitor (searches if browser URL contains this)</summary>
    public string Domain { get; set; } = string.Empty;

    /// <summary>Task name to assign when detected</summary>
    public string TaskName { get; set; } = string.Empty;
}
