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

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(System.IntPtr hWnd, out RECT lpRect);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        private delegate bool EnumWindowsProc(System.IntPtr hWnd, System.IntPtr lParam);

        // ── Win32 constants ───────────────────────────────────────────────────────

        private const uint WM_COMMAND  = 0x0111;
        private const uint BM_CLICK    = 0x00F5;
        private const int  IDOK        = 1;
        private const string DialogClass = "#32770";

        // ── State ─────────────────────────────────────────────────────────────────

        private static volatile bool _running;
        private static uint _acadPid;
        private static volatile bool _dialogLogged;  // log child controls of first dialog only

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

                // For Dialog 1 ("Reinforcement detailing"): click "Module" radio button first
                // so that Module distribution method is selected before we confirm with OK.
                // The Module CLI sequence is: start/end points → 40 (first bar offset) → 200 (spacing).
                // If the dialog is Dialog 2 ("Reinforcement description") there is no "Module"
                // child and TryClickModuleRadio is a no-op.
                // NOTE: never use IDCANCEL — it triggers WM_CANCELMODE in ASD which aborts
                //       the entire DISTRIBUTION command.
                // On first dialog: dump all child controls to debug file so we can
                // verify the Module button is found correctly.
                if (!_dialogLogged)
                {
                    _dialogLogged = true;
                    LogDialogChildren(hwnd);
                }
                TryClickModuleRadio(hwnd);
                PostMessage(hwnd, WM_COMMAND, new System.IntPtr(IDOK), System.IntPtr.Zero);

                return true;
            }, System.IntPtr.Zero);
        }

        /// <summary>
        /// Dumps all child windows of the dialog to %TEMP%\asd_dialog_children.txt for debugging.
        /// </summary>
        private static void LogDialogChildren(System.IntPtr dialogHwnd)
        {
            try
            {
                var sb = new StringBuilder();
                var titleBuf = new StringBuilder(256);
                GetWindowText(dialogHwnd, titleBuf, titleBuf.Capacity);
                sb.AppendLine($"Dialog title: '{titleBuf}'");
                sb.AppendLine("Child controls:");

                EnumChildWindows(dialogHwnd, (childHwnd, _) =>
                {
                    var cls  = new StringBuilder(64);
                    var text = new StringBuilder(128);
                    GetClassName(childHwnd, cls, cls.Capacity);
                    GetWindowText(childHwnd, text, text.Capacity);
                    RECT r2;
                    GetWindowRect(childHwnd, out r2);
                    sb.AppendLine($"  class='{cls}'  text='{text}'  L={r2.Left} T={r2.Top} R={r2.Right} B={r2.Bottom}  hwnd=0x{childHwnd.ToInt64():X}");
                    return true;
                }, System.IntPtr.Zero);

                System.IO.File.WriteAllText(
                    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "asd_dialog_children.txt"),
                    sb.ToString());
            }
            catch { }
        }

        /// <summary>
        /// Finds the "Module" distribution-method radio button in the "Reinforcement detailing"
        /// dialog and clicks it before OK is sent.
        ///
        /// ASD renders radio buttons as owner-drawn Buttons with empty text. The visible label
        /// is a separate Static control placed to the right of the radio button. Strategy:
        ///   1. Collect all child controls with their screen rects.
        ///   2. Find the Static whose text is "Module" (EN) or "Moduł" (PL).
        ///   3. Find the empty-text Button that is vertically aligned with that Static
        ///      and whose right edge is at or before the Static's left edge (radio to the left
        ///      of its label). Click it with BM_CLICK (synchronous).
        /// </summary>
        private static void TryClickModuleRadio(System.IntPtr dialogHwnd)
        {
            // Pass 1: collect all children
            var controls = new List<(System.IntPtr hwnd, string cls, string text, RECT rect)>();
            EnumChildWindows(dialogHwnd, (childHwnd, _) =>
            {
                var cls = new StringBuilder(64);
                var txt = new StringBuilder(128);
                GetClassName(childHwnd, cls, cls.Capacity);
                GetWindowText(childHwnd, txt, txt.Capacity);
                RECT r;
                GetWindowRect(childHwnd, out r);
                controls.Add((childHwnd, cls.ToString(), txt.ToString(), r));
                return true;
            }, System.IntPtr.Zero);

            // Pass 2: find the Static labelled "Module"
            RECT moduleLabelRect = default(RECT);
            bool found = false;
            foreach (var c in controls)
            {
                if (c.cls != "Static") continue;
                if (c.text.StartsWith("Module", System.StringComparison.OrdinalIgnoreCase) ||
                    c.text.StartsWith("Modu\u0142", System.StringComparison.OrdinalIgnoreCase))
                {
                    moduleLabelRect = c.rect;
                    found = true;
                    break;
                }
            }
            if (!found) return;

            int labelCenterY = (moduleLabelRect.Top + moduleLabelRect.Bottom) / 2;

            // Pass 3: find the empty Button closest to the left of the Module label
            System.IntPtr bestBtn = System.IntPtr.Zero;
            int bestScore = int.MaxValue;
            foreach (var c in controls)
            {
                if (c.cls != "Button" || c.text != "") continue;
                int btnCenterY = (c.rect.Top + c.rect.Bottom) / 2;
                int dy = System.Math.Abs(btnCenterY - labelCenterY);
                // Must be vertically close (within 15px) and to the left of the label
                if (dy > 15 || c.rect.Right > moduleLabelRect.Left + 10) continue;
                int dx = System.Math.Abs(c.rect.Right - moduleLabelRect.Left);
                int score = dy * 100 + dx;
                if (score < bestScore) { bestScore = score; bestBtn = c.hwnd; }
            }

            if (bestBtn != System.IntPtr.Zero)
                SendMessage(bestBtn, BM_CLICK, System.IntPtr.Zero, System.IntPtr.Zero);
        }
    }
}
