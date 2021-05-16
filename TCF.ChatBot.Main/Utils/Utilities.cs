using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TCF.ChatBot.Main.Utils
{
    public static class Utilities
    {
        public static List<string> ListGreetingsWords { get; set; } = new List<string>() { "OLA", "OI", "BOM DIA", "BOA TARDE", "BOA NOITE" };
        public static List<string> ListRequiredWords { get; set; } = new List<string>() { "HOW" };

        //Relative to resolution of the screen (...100% / 125% / 150%...)
        private static float increaseRate { get; set; } = 0;

        private const UInt32 MOUSEEVENTF_LEFTDOWN = 0X0002;
        private const UInt32 MOUSEEVENTF_LEFTUP = 0X0004;
        private const UInt32 MOUSEEVENTF_RIGHTDOWN = 0X0008;
        private const UInt32 MOUSEEVENTF_RIGHTUP = 0X0010;

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, uint dwExtraInf);

        [DllImport("user32.dll")]
        private static extern void SetCursorPos(int x, int y);

        [DllImport("gdi32.dll")]
        private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
        public enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117,
            //http://pinvoke.net/default.aspx/gdi32/GetDeviceCaps.html
        }

        private static float GetScalingFactor()
        {
            Graphics g = Graphics.FromHwnd(IntPtr.Zero);
            IntPtr desktop = g.GetHdc();
            int LogicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.VERTRES);
            int PhysicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.DESKTOPVERTRES);

            float ScreenScalingFactor = (float)PhysicalScreenHeight / (float)LogicalScreenHeight;

            //Get relative rate of approximation
            increaseRate = (ScreenScalingFactor - 1f) / (ScreenScalingFactor / 1f);

            return ScreenScalingFactor; // 1.25 = 125%
        }

        public static Tuple<int, int> FindExactImage(Bitmap searchImg)
        {
            var screenCapture = CaptureScreen();
            var coordinates = ContainsExact(searchImg, screenCapture);

            return coordinates;
        }

        public static Bitmap CaptureScreen()
        {
            var currentDPI = GetScalingFactor();
            var width = (int)(SystemInformation.VirtualScreen.Width * currentDPI);
            var height = (int)(SystemInformation.VirtualScreen.Height * currentDPI);
            Bitmap screenCapture = new Bitmap(width, height);
            Graphics graphic = Graphics.FromImage(screenCapture);

            graphic.CopyFromScreen(Screen.PrimaryScreen.Bounds.X,
                             Screen.PrimaryScreen.Bounds.Y,
                             0, 0,
                             screenCapture.Size,
                             CopyPixelOperation.SourceCopy);
#if DEBUG
            screenCapture.Save($@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\printScreen.png");
#endif
            return screenCapture;
        }

        private static Tuple<int, int> ContainsExact(Bitmap searchImg, Bitmap sourceImg)
        {
            for (int x = 0; x < sourceImg.Width; x++)
            {
                for (int y = 0; y < sourceImg.Height; y++)
                {
                    bool invalid = false;
                    int k = x, l = y;
                    for (int a = 0; a < searchImg.Width; a++)
                    {
                        l = y;
                        for (int b = 0; b < searchImg.Height; b++)
                        {
                            if (searchImg.GetPixel(a, b) != sourceImg.GetPixel(k, l))
                            {
                                invalid = true;
                                break;
                            }
                            else
                            {
                                //Click(k, l);
                                l++;
                            }
                        }

                        if (invalid)
                            break;
                        else //Probably it has found
                            k++;
                    }

                    if (!invalid)
                    {
                        var coordinateX = (k - searchImg.Width / 2);
                        var coordinateY = (l - searchImg.Width / 2);
                        return new Tuple<int, int>(coordinateX, coordinateY);
                    }
                }
            }
            return new Tuple<int, int>(-1, -1);
        }

        public static Tuple<int, int> SearchColor(string hexColor, int width, int height, int coordinateX = 0, int coordinateY = 0, bool setMousePosition = false)
        {
            //Instance of a new object
            Tuple<int, int> coordinates = new Tuple<int, int>(-1, -1);

            try
            {
                var winWidth = width;
                var winHeight = height;
                var winPosX = coordinateX;
                var winPosY = coordinateY;

                //Percentage of layout scale (125%, 100%, etc)

                //Create an empty bitmap with the size of all connected screens
                Bitmap bitmap = new Bitmap(winWidth, winHeight);

                //Create a new graphic object that can capture the screen
                Graphics graphics = Graphics.FromImage(bitmap as Image);

                //Screenshot moment > screen content to graphics object
                graphics.CopyFromScreen(winPosX, winPosY, 0, 0, bitmap.Size);

#if DEBUG
                bitmap.Save($@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\screen.png");
#endif
                //e.g translate #ffffff to a Color object
                Color desiredPixelColor = ColorTranslator.FromHtml(hexColor);

                for (int x = 0; x < winWidth; x++)
                {
                    for (int y = 0; y < winHeight; y++)
                    {
                        if (setMousePosition)
                            Click(x + winPosX, y + winPosY);

                        //Get current pixels color
                        Color currentPixelColor = bitmap.GetPixel(x, y);
                        //Console.WriteLine(currentPixelColor.Name);
                        if (desiredPixelColor.Name == currentPixelColor.Name)
                        {
                            //MessageBox.Show($"Found pixel at {x}, {y}");
                            coordinates = new Tuple<int, int>(x + winPosX, y + winPosY);
                            return coordinates;
                        }
                    }
                }
            }
            catch (Exception ex)
            { }

            return coordinates;
        }

        public static int GetEquivalent(int val)
        {
            return decimal.ToInt32(val - ((val) * (decimal)increaseRate));
        }

        public static void Click(int x, int y, int times = 1)
        {
            SetCursorPos(GetEquivalent(x), GetEquivalent(y));

            for (int i = 0; i < times; i++)
            {
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
            }
        }

        public static void RightClick(int x, int y, int times = 1)
        {
            SetCursorPos(GetEquivalent(x), GetEquivalent(y));

            for (int i = 0; i < times; i++)
            {
                mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
                mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
            }
        }

        public static void SendCtrlC()
        {
            Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
            SendKeys.SendWait("^(c)");
            Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
        }

        public static void SendCtrlV()
        {
            Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
            SendKeys.SendWait("^(v)");
            Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
        }

        public static void SendCtrlA()
        {
            Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
            SendKeys.SendWait("^(a)");
            Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
        }

        public static void SendEnter()
        {
            Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
            SendKeys.SendWait("{ENTER}");
            Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
        }

        public static void SendShiftEnter()
        {
            Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
            SendKeys.SendWait("+{ENTER}");
            Task.Delay(TimeSpan.FromMilliseconds(10)).Wait();
        }

        public static void SendText(string text)
        {
            SendKeys.SendWait(text);
        }

        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }
    }
}
