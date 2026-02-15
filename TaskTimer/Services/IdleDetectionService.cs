using System.Runtime.InteropServices;

namespace TaskTimer.Services;

/// <summary>
/// マウス/キーボードのアイドル時間を検出するサービス（Win32 API使用）
/// </summary>
public class IdleDetectionService : IDisposable
{
    [DllImport("user32.dll")]
    private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

    [StructLayout(LayoutKind.Sequential)]
    private struct LASTINPUTINFO
    {
        public uint cbSize;
        public uint dwTime;
    }

    private readonly System.Windows.Threading.DispatcherTimer _timer;
    private readonly int _idleThresholdMs;
    private bool _isIdle;

    public event EventHandler? IdleStarted;
    public event EventHandler? IdleEnded;

    public bool IsIdle => _isIdle;

    public IdleDetectionService(int idleThresholdSeconds)
    {
        _idleThresholdMs = idleThresholdSeconds * 1000;
        _timer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1)
        };
        _timer.Tick += CheckIdleState;
    }

    public void Start() => _timer.Start();
    public void Stop() => _timer.Stop();

    private void CheckIdleState(object? sender, EventArgs e)
    {
        var idleTime = GetIdleTimeMs();
        if (idleTime >= _idleThresholdMs && !_isIdle)
        {
            _isIdle = true;
            IdleStarted?.Invoke(this, EventArgs.Empty);
        }
        else if (idleTime < _idleThresholdMs && _isIdle)
        {
            _isIdle = false;
            IdleEnded?.Invoke(this, EventArgs.Empty);
        }
    }

    private static uint GetIdleTimeMs()
    {
        var info = new LASTINPUTINFO { cbSize = (uint)Marshal.SizeOf<LASTINPUTINFO>() };
        if (GetLastInputInfo(ref info))
        {
            return (uint)Environment.TickCount - info.dwTime;
        }
        return 0;
    }

    public void Dispose()
    {
        _timer.Stop();
        GC.SuppressFinalize(this);
    }
}
