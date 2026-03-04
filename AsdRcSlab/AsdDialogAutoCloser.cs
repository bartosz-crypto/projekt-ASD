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
    ///   "Reinforcement detailing"   → OK  (accepts current Linear/Module/Top settings)
    ///   "Reinforcement description" → OK  (saves default description, bars are placed)
    ///
    /// Using IDCANCEL on the description dialog sends WM_CANCELMODE to AutoCAD,
    /// which aborts the distribution command — so we always use IDOK.
    ///
    /// No keyboard events are sent — avoids any key propagation to AutoCAD.
    /// Polls every 80 ms. Filters by process ID and window class.
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
        private static extern System.IntPtr SendMessage(System.IntPtr hWnd, uint Msg,
            System.IntPtr wParam, System.IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetClassName(System.IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(System.IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(System.IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        private delegate bool EnumWindowsProc(System.IntPtr hWnd, System.IntPtr lParam);

        // ── Win32 constants ───────────────────────────────────────────────────────

        private const uint WM_COMMAND  = 0x0111;
        private const uint BM_CLICK    = 0x00F5;
        private const int  IDOK        = 1;
        private const string DialogClass = "#32770";

        // ── State ─────────────────────────────────────────────────────────────────

        private static volatile bool _running;
        private static uint _acadPid;

        // ── Public API ────────────────────────────────────────────────────────────

        /// <summary>
        /// Start monitoring. Call just before SendStringToExecute.
        /// Stops automatically after <paramref name="timeoutMs"/> ms.
        /// </summary>
        public static void Start(int timeoutMs = 300_000)
        {
            _acadPid = (uint)Process.GetCurrentProcess().Id;
            _running = true;
            Task.Run(Monitor);
            Task.Delay(timeoutMs).ContinueWith(_ => Stop());
        }

        public static void Stop() => _running = false;

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

                // Only AutoCAD's own windows
                uint pid;
                GetWindowThreadProcessId(hwnd, out pid);
                if (pid != _acadPid) return true;

                // Only Win32 dialogs
                var cls = new StringBuilder(64);
                GetClassName(hwnd, cls, cls.Capacity);
                if (cls.ToString() != DialogClass) return true;

                // Close with OK — handles ASD dialogs in any language (English, Polish, etc.)
                // NOTE: TryClickModuleRadio disabled — was causing unpredictable dialog behaviour.
                // Zone distribution with spacing=lenKey also places exactly 1 bar, so method
                // selection is not critical for correct bar count.
                // Both ASD dialogs are closed with IDOK:
                //   Dialog 1 (distribution settings) → OK accepts the current settings
                //   Dialog 2 (bar description)       → OK saves default description
                // NOTE: never use IDCANCEL — it triggers WM_CANCELMODE in ASD which aborts
                //       the entire DISTRIBUTION command.
                PostMessage(hwnd, WM_COMMAND, new System.IntPtr(IDOK), System.IntPtr.Zero);

                return true;
            }, System.IntPtr.Zero);
        }

        /// <summary>
        /// Searches child controls of <paramref name="dialogHwnd"/> for a button/radio whose
        /// text is "Module" (or starts with "Module"/"Moduł" for Polish installs).
        /// Sends BM_CLICK to it so the Module distribution method is selected before OK.
        /// </summary>
        private static void TryClickModuleRadio(System.IntPtr dialogHwnd)
        {
            EnumChildWindows(dialogHwnd, (childHwnd, _) =>
            {
                var text = new StringBuilder(64);
                GetWindowText(childHwnd, text, text.Capacity);
                var t = text.ToString();

                // Match English "Module" or Polish "Moduł" (or any prefix variant)
                if (t.StartsWith("Module", System.StringComparison.OrdinalIgnoreCase) ||
                    t.StartsWith("Modu\u0142", System.StringComparison.OrdinalIgnoreCase))
                {
                    // SendMessage is synchronous — the click is processed before we proceed to IDOK
                    SendMessage(childHwnd, BM_CLICK, System.IntPtr.Zero, System.IntPtr.Zero);
                    return false; // stop child enumeration
                }
                return true;
            }, System.IntPtr.Zero);
        }
    }
}
