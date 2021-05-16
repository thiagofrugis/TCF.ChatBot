using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TCF.ChatBot.Main.Utils
{
    public class InterceptKeyboard
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_KEYUP = 0x101;
        private const int WM_SYSKEYUP = 0x105;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        public event EventHandler<Keys> OnKeyPressed;
        public event EventHandler<Keys> OnKeyUnpressed;

        private LowLevelKeyboardProc proc;
        private IntPtr hookID = IntPtr.Zero;

        public InterceptKeyboard()
        {
            proc = HookCallback;
        }

        public void HookKeyboard()
        {
            hookID = SetHook(proc);
        }

        public void UnHookKeyboard()
        {
            UnhookWindowsHookEx(hookID);
        }

        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                OnKeyPressed.Invoke(this, ((Keys)vkCode));
            }
            else if (nCode >= 0 && wParam == (IntPtr)WM_KEYUP || wParam == (IntPtr)WM_SYSKEYUP)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                OnKeyUnpressed.Invoke(this, ((Keys)vkCode));
            }

            return CallNextHookEx(hookID, nCode, wParam, lParam);
        }

        public static void SetInterceptKeyboard(Keys key)
        {
            //MustBeAllPressed = false;
            KeysPressed = new Dictionary<string, bool>();

            KeysPressed.Add(key.ToString(), false);

            var ik = new InterceptKeyboard();
            ik.OnKeyPressed += InterceptKeyboard.InterceptOnKeyPressed;
            ik.OnKeyUnpressed += InterceptKeyboard.InterceptOnKeyUnpressed;
            ik.HookKeyboard();
        }

        public static void SetInterceptKeyboard(bool mustAllBePressed = true, params Keys[] keys)
        {
            MustAllBePressed = mustAllBePressed;
            KeysPressed = new Dictionary<string, bool>();

            foreach (var key in keys)
                KeysPressed.Add(key.ToString(), false);

            var ik = new InterceptKeyboard();
            ik.OnKeyPressed += InterceptKeyboard.InterceptOnKeyPressed;
            ik.OnKeyUnpressed += InterceptKeyboard.InterceptOnKeyUnpressed;
            ik.HookKeyboard();
        }

        private static Dictionary<string, bool> KeysPressed;
        private static bool MustAllBePressed;
        public static bool stopExecuting { get; set; }

        private static void InterceptOnKeyPressed(object sender, Keys e)
        {
            //Pressed preset key
            if (KeysPressed.ContainsKey(e.ToString()))
                KeysPressed[e.ToString()] = true;

            CheckKeyBehavior();
        }

        private static void InterceptOnKeyUnpressed(object sender, Keys e)
        {
            //Unpressed preset key
            if (KeysPressed.ContainsKey(e.ToString()))
                KeysPressed[e.ToString()] = false;
        }

        private static void CheckKeyBehavior()
        {
            //If any or all keys are pressed, it will close application
            if ((MustAllBePressed && KeysPressed.All(_ => _.Value)) ||
                (!MustAllBePressed && KeysPressed.Any(_ => _.Value)))
            {
                stopExecuting = true;
                Environment.Exit(0);
            }
        }
    }
}