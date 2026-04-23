using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.IO;
using System.Collections.Generic;

namespace Workey3101
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")] static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")] static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] static extern short GetAsyncKeyState(int vKey);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)] static extern IntPtr GetModuleHandle(string lpModuleName);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")] static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr lpdwProcessId);
        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("imm32.dll")] static extern IntPtr ImmGetContext(IntPtr hWnd);
        [DllImport("imm32.dll")] static extern bool ImmSetOpenStatus(IntPtr hIMC, bool fOpen);
        [DllImport("imm32.dll")] static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);
        [DllImport("imm32.dll")] static extern bool ImmNotifyIME(IntPtr hIMC, uint dwAction, uint dwIndex, uint dwValue);
        [DllImport("imm32.dll")] static extern bool ImmSetCompositionWindow(IntPtr hIMC, ref COMPOSITIONFORM lpCompForm);

        [StructLayout(LayoutKind.Sequential)]
        struct COMPOSITIONFORM { public uint dwStyle; public POINT ptCurrentPos; public RECT rcArea; }
        [StructLayout(LayoutKind.Sequential)]
        struct POINT { public int x; public int y; }
        [StructLayout(LayoutKind.Sequential)]
        struct RECT { public int left; public int top; public int right; public int bottom; }
        const uint CFS_POINT = 0x0002;
        const uint NI_COMPOSITIONSTR = 0x0015;
        const uint CPS_COMPLETE = 0x0001;
        const int WM_PASTE = 0x0302;

        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        const uint SWP_NOMOVE = 0x0002, SWP_NOSIZE = 0x0001, SWP_SHOWWINDOW = 0x0040;

        const int WH_KEYBOARD_LL = 13;
        const int WM_KEYDOWN = 0x0100, WM_SYSKEYDOWN = 0x0104;
        const int VK_CONTROL = 0x11, VK_MENU = 0x12, VK_APPS = 0x5D, VK_SNAPSHOT = 0x2C;

        [StructLayout(LayoutKind.Sequential)] struct KEYBDINPUT { public ushort wVk, wScan; public uint dwFlags, time; public IntPtr dwExtraInfo; }
        [StructLayout(LayoutKind.Sequential)] struct MOUSEINPUT { public int dx, dy; public uint mouseData, dwFlags, time; public IntPtr dwExtraInfo; }
        [StructLayout(LayoutKind.Explicit)]
        struct INPUT
        {
            [FieldOffset(0)] public uint type;
            [FieldOffset(4)] public KEYBDINPUT ki;
            [FieldOffset(4)] public MOUSEINPUT mi;
        }
        const uint INPUT_KEYBOARD = 1, KEYEVENTF_KEYUP = 0x0002;

        delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        [StructLayout(LayoutKind.Sequential)]
        struct KBDLLHOOKSTRUCT { public uint vkCode, scanCode, flags, time; public IntPtr dwExtraInfo; }

        LowLevelKeyboardProc _hookProc;
        IntPtr _hookId = IntPtr.Zero;
        bool _on = false;
        bool _windowVisible = false;
        bool _initialized = false;
        bool _sending = false;
        bool _windowOpening = false;
        IntPtr _prevWindow = IntPtr.Zero;
        PendingKey _triggerKey;
        bool _hasTriggerKey = false;
        int _fontSize = 21;
        Queue<PendingKey> _pendingKeys = new Queue<PendingKey>();

        struct PendingKey { public uint vk; public bool shift; }

        RichTextBox _textBox;
        NotifyIcon _trayIcon;

        const int WIN_H = 60;

        public Form1()
        {
            InitializeComponent();
            this.Opacity = 0;
            this.ShowInTaskbar = false;
            BuildUI();
            InstallHook();
            this.Hide();
        }

        void BuildUI()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.TopMost = true;
            this.BackColor = Color.White;
            this.ShowInTaskbar = false;
            this.AutoScaleMode = AutoScaleMode.None;
            this.Size = new Size(100, WIN_H);

            var outer = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            this.Controls.Add(outer);

            _textBox = new RichTextBox
            {
                Dock = DockStyle.Bottom,
                Height = 35,
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, _fontSize),
                BackColor = Color.White,
                ForeColor = Color.Black,
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.None,
                Multiline = false,
                WordWrap = false,
            };
            outer.Controls.Add(_textBox);
            _textBox.KeyDown += TextBox_KeyDown;

            this.Deactivate += (s, e) => {
                if (_windowVisible && _initialized)
                    this.BeginInvoke((Action)(() => {
                        if (!string.IsNullOrWhiteSpace(_textBox.Text))
                            SendAndCloseNoFocus();
                        else
                            HideInputWindow();
                    }));
            };

            var initTimer = new System.Windows.Forms.Timer { Interval = 3000 };
            initTimer.Tick += (s, e) => { _initialized = true; initTimer.Stop(); };
            initTimer.Start();

            var imeTimer = new System.Windows.Forms.Timer { Interval = 200 };
            imeTimer.Tick += (s, e) => { if (_windowVisible) ImeOn(); };
            imeTimer.Start();

            Application.ApplicationExit += (s, e) => CleanUpTrayIcon();

            _trayIcon = new NotifyIcon { Visible = true, Text = "Workey OFF", Icon = CreateTextIcon("英", Color.Gray, Color.White) };
            _trayIcon.MouseClick += (s, e) => {
                if (e.Button == MouseButtons.Left)
                {
                    if (_on) SetOff(); else SetOn();
                }
            };
            var menu = new ContextMenuStrip();
            menu.Items.Add("終了", null, (s, e) => {
                if (_windowVisible && !string.IsNullOrWhiteSpace(_textBox.Text)) SendAndCloseNoFocus();
                else HideInputWindow();
                CleanUpTrayIcon();
                Application.Exit();
            });
            _trayIcon.ContextMenuStrip = menu;
            
            SetOn();
        }

        void SetOn()
        {
            _on = true;
            if (_trayIcon.Icon != null)
            {
                _trayIcon.Icon.Dispose();
                _trayIcon.Icon = null;
            }
            _trayIcon.Icon = CreateTextIcon("全", Color.White, Color.Red);
            _trayIcon.Text = "Workey ON";
            _trayIcon.BalloonTipTitle = "Workey3101";
            _trayIcon.BalloonTipText = "ON になりました";
            _trayIcon.ShowBalloonTip(1500);
        }

        void SetOff()
        {
            _on = false;
            if (_trayIcon.Icon != null)
            {
                _trayIcon.Icon.Dispose();
                _trayIcon.Icon = null;
            }
            _trayIcon.Icon = CreateTextIcon("英", Color.Gray, Color.White);
            _trayIcon.Text = "Workey OFF";
            _trayIcon.BalloonTipTitle = "Workey3101";
            _trayIcon.BalloonTipText = "OFF になりました";
            _trayIcon.ShowBalloonTip(1500);
        }

        void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                e.SuppressKeyPress = true;
                if (string.IsNullOrWhiteSpace(_textBox.Text)) HideInputWindow();
                else SendAndClose();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                _textBox.Clear();
                HideInputWindow();
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                e.SuppressKeyPress = true;
                if (Clipboard.ContainsText()) _textBox.SelectedText = Clipboard.GetText();
            }
            else if (e.KeyCode == Keys.Back && string.IsNullOrEmpty(_textBox.Text))
            {
                e.SuppressKeyPress = true;
                HideInputWindow();
            }
        }

        void InstallHook()
        {
            _hookProc = HookCallback;
            _hookId = SetWindowsHookEx(WH_KEYBOARD_LL, _hookProc, GetModuleHandle(null), 0);
        }

        void UninstallHook()
        {
            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }

        string GetForegroundClass()
        {
            var fw = GetForegroundWindow();
            if (fw == IntPtr.Zero || fw == this.Handle) return "";
            var sb = new System.Text.StringBuilder(256);
            GetClassName(fw, sb, 256);
            return sb.ToString();
        }

        IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            var kbPre = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);

            // Ctrl+メニューキーでON/OFF
            if (kbPre.vkCode == VK_APPS)
            {
                if ((int)wParam == WM_KEYDOWN || (int)wParam == WM_SYSKEYDOWN)
                {
                    if ((GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0)
                    {
                        this.Invoke((Action)(() => {
                            if (_on)
                            {
                                if (_windowVisible && !string.IsNullOrWhiteSpace(_textBox.Text)) SendAndCloseNoFocus();
                                else HideInputWindow();
                                SetOff();
                            }
                            else SetOn();
                        }));
                    }
                }
                return (IntPtr)1;
            }

            if (nCode >= 0)
            {
                var kb = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                uint vk = kb.vkCode;
                int msg = (int)wParam;
                bool down = msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN;
                bool ctrl = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
                bool alt = (GetAsyncKeyState(VK_MENU) & 0x8000) != 0;
                bool shift = (GetAsyncKeyState(0x10) & 0x8000) != 0;

                // 窓内で半角全角・ひらがなカタカナ無効化
                if (_windowVisible && (vk == 0xF2 || vk == 0xF3 || vk == 0xF4 || vk == 0x15))
                    return (IntPtr)1;

                bool isTypingKey =
                    (vk >= 0x41 && vk <= 0x5A) ||  // A-Z
                    vk == 0xBC ||                    // , → 、
                    vk == 0xBE ||                    // . → 。
                    vk == 0xDB ||                    // [ → 「
                    vk == 0xDD ||                    // ] → 」
                    vk == 0xDE;                      // '

                // 窓が開く途中のキーを貯める
                if (down && _windowOpening && !ctrl && !alt && isTypingKey)
                {
                    _pendingKeys.Enqueue(new PendingKey { vk = vk, shift = shift });
                    return (IntPtr)1;
                }

                // Word（OpusApp）のときだけ窓を開く
                if (down && _on && !_windowVisible && !_windowOpening && !ctrl && !alt && isTypingKey)
                {
                    string cls = GetForegroundClass();
                    var sbT = new System.Text.StringBuilder(256);
                    GetWindowText(GetForegroundWindow(), sbT, 256);
                    string title = sbT.ToString();
                    bool isWord = cls == "OpusApp" ||
                                  title.EndsWith("- Word") ||
                                  title.Contains(".docx");
                    if (!isWord) return CallNextHookEx(_hookId, nCode, wParam, lParam);
                    _triggerKey = new PendingKey { vk = vk, shift = shift };
                    _hasTriggerKey = true;
                    _windowOpening = true;
                    this.Invoke((Action)ShowInputWindow);
                    return (IntPtr)1;
                }
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        void ShowInputWindow()
        {
            var fw = GetForegroundWindow();
            if (fw != IntPtr.Zero && fw != this.Handle)
                _prevWindow = fw;

            RECT wr; GetWindowRect(_prevWindow, out wr);
            var pt = new System.Drawing.Point(wr.left + (wr.right - wr.left) / 2, wr.top + (wr.bottom - wr.top) / 2);
            var screen = Screen.FromPoint(pt).Bounds;
            int winW = screen.Width / 2;
            int x = screen.Left + (screen.Width - winW) / 2;
            int y = screen.Top + screen.Height - WIN_H;
            this.Size = new Size(winW, WIN_H);
            this.Location = new Point(x, y);

            _windowVisible = true; this.Opacity = 1;
            SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            this.Show(); ForceFocus();
            ImeOn(); SetCandidatePosition();
            Clipboard.Clear();

            if (_hasTriggerKey)
            {
                var trigger = _triggerKey;
                _hasTriggerKey = false;
                var pending = new Queue<PendingKey>(_pendingKeys);
                _pendingKeys.Clear();

                var t = new System.Windows.Forms.Timer { Interval = 200 };
                t.Tick += (s, e) => {
                    t.Stop();
                    _windowOpening = false;
                    AppendPendingKeysToTextBox(trigger, pending);
                };
                t.Start();
            }
            else
            {
                _windowOpening = false;
                _pendingKeys.Clear();
            }
        }

        void AppendPendingKeysToTextBox(PendingKey trigger, Queue<PendingKey> pending)
        {
            _textBox.Focus();
            UninstallHook();

            var inputs = new List<INPUT>();
            AddKeyInput(trigger, inputs);
            foreach (var pk in pending)
                AddKeyInput(pk, inputs);

            if (inputs.Count > 0)
                SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf(typeof(INPUT)));

            InstallHook();
        }

        void AddKeyInput(PendingKey pk, List<INPUT> inputs)
        {
            if (pk.shift)
                inputs.Add(new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = 0x10, wScan = 0, dwFlags = 0, time = 0, dwExtraInfo = IntPtr.Zero } });

            inputs.Add(new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = (ushort)pk.vk, wScan = 0, dwFlags = 0, time = 0, dwExtraInfo = IntPtr.Zero } });
            inputs.Add(new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = (ushort)pk.vk, wScan = 0, dwFlags = KEYEVENTF_KEYUP, time = 0, dwExtraInfo = IntPtr.Zero } });

            if (pk.shift)
                inputs.Add(new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = 0x10, wScan = 0, dwFlags = KEYEVENTF_KEYUP, time = 0, dwExtraInfo = IntPtr.Zero } });
        }

        void HideInputWindow() { _windowVisible = false; this.Opacity = 0; this.Hide(); }

        void ForceFocus()
        {
            try
            {
                var fw = GetForegroundWindow();
                uint fwt = GetWindowThreadProcessId(fw, IntPtr.Zero), myt = GetCurrentThreadId();
                AttachThreadInput(fwt, myt, true);
                SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                SetForegroundWindow(this.Handle); _textBox.Focus();
                AttachThreadInput(fwt, myt, false);
            }
            catch { }
        }

        void ImeOn()
        {
            var himc = ImmGetContext(_textBox.Handle);
            if (himc != IntPtr.Zero) { ImmSetOpenStatus(himc, true); ImmReleaseContext(_textBox.Handle, himc); }
        }

        void ImeForceCommit()
        {
            var himc = ImmGetContext(_textBox.Handle);
            if (himc != IntPtr.Zero)
            {
                ImmNotifyIME(himc, NI_COMPOSITIONSTR, CPS_COMPLETE, 0);
                ImmReleaseContext(_textBox.Handle, himc);
            }
        }

        void SetCandidatePosition()
        {
            var himc = ImmGetContext(_textBox.Handle);
            if (himc != IntPtr.Zero)
            {
                var cf = new COMPOSITIONFORM { dwStyle = CFS_POINT, ptCurrentPos = new POINT { x = 500, y = 4 } };
                ImmSetCompositionWindow(himc, ref cf);
                ImmReleaseContext(_textBox.Handle, himc);
            }
        }

        void SendAndClose()
        {
            if (_sending) return;
            _sending = true;
            ImeForceCommit();
            Thread.Sleep(50);
            string text = _textBox.Text.Trim();
            IntPtr target = _prevWindow;
            HideInputWindow();
            _textBox.Clear();
            if (!string.IsNullOrEmpty(text) && target != IntPtr.Zero)
            {
                Clipboard.SetText(text);
                SetForegroundWindow(target);
                Thread.Sleep(150);
                SendCtrlV();
            }
            _sending = false;
        }

        void SendAndCloseNoFocus()
        {
            if (_sending) return;
            _sending = true;
            ImeForceCommit();
            Thread.Sleep(50);
            string text = _textBox.Text.Trim();
            IntPtr target = _prevWindow;
            HideInputWindow();
            _textBox.Clear();
            if (!string.IsNullOrEmpty(text) && target != IntPtr.Zero)
            {
                Clipboard.SetText(text);
                SetForegroundWindow(target);
                Thread.Sleep(100);
                SendCtrlV();
            }
            _sending = false;
        }

        void SendCtrlV()
        {
            if (_prevWindow == IntPtr.Zero) return;

            SetForegroundWindow(_prevWindow);
            Thread.Sleep(50);

            var pasteInputs = new[]
            {
                new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = 0x11, wScan = 0, dwFlags = 0, time = 0, dwExtraInfo = IntPtr.Zero } },
                new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = 0x56, wScan = 0, dwFlags = 0, time = 0, dwExtraInfo = IntPtr.Zero } },
                new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = 0x56, wScan = 0, dwFlags = KEYEVENTF_KEYUP, time = 0, dwExtraInfo = IntPtr.Zero } },
                new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = 0x11, wScan = 0, dwFlags = KEYEVENTF_KEYUP, time = 0, dwExtraInfo = IntPtr.Zero } },
            };

            SendInput((uint)pasteInputs.Length, pasteInputs, Marshal.SizeOf(typeof(INPUT)));
            Thread.Sleep(50);

            var enterInputs = new[]
            {
                new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = 0x0D, wScan = 0, dwFlags = 0, time = 0, dwExtraInfo = IntPtr.Zero } },
                new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = 0x0D, wScan = 0, dwFlags = KEYEVENTF_KEYUP, time = 0, dwExtraInfo = IntPtr.Zero } },
            };

            SendInput((uint)enterInputs.Length, enterInputs, Marshal.SizeOf(typeof(INPUT)));
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool DestroyIcon(IntPtr hIcon);

        Icon CreateTextIcon(string text, Color foreColor, Color backColor)
        {
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(backColor);
                using (var font = new Font("Arial", 8f, FontStyle.Bold))
                using (var brush = new SolidBrush(foreColor))
                    g.DrawString(text, font, brush, new RectangleF(0, 0, 16, 16),
                        new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
            }
            var hicon = bmp.GetHicon();
            bmp.Dispose(); // メモリーリーク防止
            if (hicon == IntPtr.Zero) return SystemIcons.Application;

            Icon icon;
            using (var tempIcon = Icon.FromHandle(hicon))
            {
                icon = (Icon)tempIcon.Clone();
            }
            DestroyIcon(hicon);
            return icon;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            UninstallHook();
            CleanUpTrayIcon();
            base.OnFormClosing(e);
        }

        void CleanUpTrayIcon()
        {
            if (_trayIcon != null)
            {
                try
                {
                    _trayIcon.Visible = false;
                    if (_trayIcon.Icon != null)
                    {
                        _trayIcon.Icon.Dispose();
                        _trayIcon.Icon = null;
                    }
                    if (_trayIcon.ContextMenuStrip != null)
                    {
                        _trayIcon.ContextMenuStrip.Dispose();
                        _trayIcon.ContextMenuStrip = null;
                    }
                    _trayIcon.Dispose();
                }
                catch { }
                finally
                {
                    _trayIcon = null;
                }
            }
        }
    }
}