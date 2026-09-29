using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AsdRcSlab
{
    /// <summary>
    /// Monitors for ASD modal dialogs in AutoCAD's process and automatically
    /// confirms them with OK while a batch DISTRIBUTION command is executing.
    ///
    /// IDOK is sent to BOTH ASD dialogs:
    ///   "Reinforcement detailing"   → OK  (accepts current method settings)
    ///   "Reinforcement description" → OK  (saves default description, bars are placed)
    ///
    /// No Module radio click: let ASD use whatever method is currently selected.
    /// Using IDCANCEL on ANY dialog sends WM_CANCELMODE → aborts DISTRIBUTION.
    /// Polls every 80 ms. Filters by process ID and window class.
    /// Logs all activity to %TEMP%\asd_closers_log.txt for diagnosis.
    /// </summary>
    internal static class AsdDialogAutoCloser
    {
        // ── P/Invoke ─────────────────────────────────────────────────────────────

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, System.IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(System.IntPtr hWnd, EnumWindowsProc lpEnumFunc, System.IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(System.IntPtr hWnd, uint Msg,
            System.IntPtr wParam, System.IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetClassName(System.IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(System.IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(System.IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(System.IntPtr hWnd, out RECT lpRect);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        private delegate bool EnumWindowsProc(System.IntPtr hWnd, System.IntPtr lParam);

        // ── Win32 constants ───────────────────────────────────────────────────────

        private const uint WM_COMMAND  = 0x0111;
        private const int  IDOK        = 1;
        private const string DialogClass = "#32770";

        // ── State ─────────────────────────────────────────────────────────────────

        private static volatile bool _running;
        private static uint _acadPid;

        // Tracks HWNDs recently dismissed (time-based expiry 500 ms) to avoid
        // double-dismissing the same dialog before it closes.
        private static readonly Dictionary<System.IntPtr, int> _processedHwnds =
            new Dictionary<System.IntPtr, int>();
        private const int ProcessedHwndExpiryMs = 500;

        // ── Logging ───────────────────────────────────────────────────────────────

        private static readonly string _logPath =
            System.IO.Path.Combine(System.IO.Path.GetTempPath(), "asd_closers_log.txt");

        private static void Log(string msg)
        {
            try { System.IO.File.AppendAllText(_logPath,
                $"[{System.DateTime.Now:HH:mm:ss.fff}] {msg}\r\n"); }
            catch { }
        }

        // ── Public API ────────────────────────────────────────────────────────────

        public static void Start(int timeoutMs = 300_000)
        {
            _acadPid = (uint)Process.GetCurrentProcess().Id;
            _processedHwnds.Clear();
            _running = true;
            try { System.IO.File.WriteAllText(_logPath,
                $"[{System.DateTime.Now:HH:mm:ss.fff}] === START pid={_acadPid} ===\r\n"); }
            catch { }
            Task.Run(Monitor);
            Task.Delay(timeoutMs).ContinueWith(_ => Stop());
        }

        public static void Stop()
        {
            _running = false;
            Log("=== STOP ===");
        }

        // ── Monitor loop ──────────────────────────────────────────────────────────

        private static void Monitor()
        {
            while (_running)
            {
                Thread.Sleep(80);
                CloseAsdDialogs();
            }
        }

        private static void CloseAsdDialogs()
        {
            EnumWindows((hwnd, _) =>
            {
                if (!IsWindowVisible(hwnd)) return true;

                uint pid;
                GetWindowThreadProcessId(hwnd, out pid);
                if (pid != _acadPid) return true;

                var cls = new StringBuilder(64);
                GetClassName(hwnd, cls, cls.Capacity);
                if (cls.ToString() != DialogClass) return true;

                // Skip recently-dismissed HWNDs (PostMessage async → dialog still visible)
                if (_processedHwnds.TryGetValue(hwnd, out int closedAt) &&
                    unchecked(System.Environment.TickCount - closedAt) < ProcessedHwndExpiryMs)
                    return true;

                // Read title + child text for logging
                var titleBuf = new StringBuilder(256);
                GetWindowText(hwnd, titleBuf, titleBuf.Capacity);
                string title = titleBuf.ToString();

                var childTxt = new StringBuilder();
                EnumChildWindows(hwnd, (ch, lp) =>
                {
                    var t = new StringBuilder(256);
                    GetWindowText(ch, t, t.Capacity);
                    if (t.Length > 0) childTxt.Append('[').Append(t).Append(']');
                    return true;
                }, System.IntPtr.Zero);

                // Dismiss with IDOK — works for all ASD dialogs in any language.
                // NOTE: never use IDCANCEL → triggers WM_CANCELMODE → aborts DISTRIBUTION.
                PostMessage(hwnd, WM_COMMAND, new System.IntPtr(IDOK), System.IntPtr.Zero);
                _processedHwnds[hwnd] = System.Environment.TickCount;

                Log($"IDOK → hwnd=0x{hwnd.ToInt64():X} title='{title}' | {childTxt}");

                return true;
            }, System.IntPtr.Zero);
        }
    }
}
