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
using TCF.ChatBot.Main.Domain;

namespace TCF.ChatBot.Main.Utils
{
    public static class Utilities
    {
        private static string GreenColor { get; } = "#06d756";
        private static string GreenColor2 { get; } = "#09d757";
        private static string BottomColor { get; } = "#f0f0f0";
        private static string GreyColor1 { get; } = "#919191";
        private static string GreyColor2 { get; } = "#919191";
        private static decimal windowLayoutRate { get; } = 120;
        private static decimal increaseRate { get; set; } = (windowLayoutRate / 100) - 1;

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
            var width = (int)(Screen.PrimaryScreen.Bounds.Width * currentDPI);
            var height = (int)(Screen.PrimaryScreen.Bounds.Height * currentDPI);
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

        //        public static Coordinates SearchGreenButton()
        //        {
        //            //Instance of a new object
        //            Coordinates coordinates = new Coordinates();

        //            try
        //            {
        //                var winWidth = SystemInformation.VirtualScreen.Width;
        //                var winHeight = SystemInformation.VirtualScreen.Height;
        //                var winPosX = 78;
        //                var winPosY = 0;

        //                //Percentage of layout scale (125%, 100%, etc)

        //                //Create an empty bitmap with the size of all connected screens
        //                Bitmap bitmap = new Bitmap(winWidth, winHeight);

        //                //Create a new graphic object that can capture the screen
        //                Graphics graphics = Graphics.FromImage(bitmap as Image);

        //                //Screenshot moment > screen content to graphics object
        //                graphics.CopyFromScreen(winPosX, winPosY, 0, 0, bitmap.Size);

        //#if DEBUG
        //                bitmap.Save($@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\screen.png");
        //#endif
        //                //e.g translate #ffffff to a Color object
        //                Color desiredPixelColor = ColorTranslator.FromHtml(GreenColor);
        //                Color desiredPixelColor2 = ColorTranslator.FromHtml(GreenColor2);

        //                for (int x = 0; x < winWidth; x++)
        //                {
        //                    for (int y = 0; y < winHeight; y++)
        //                    {
        //                        //Get current pixels color
        //                        Color currentPixelColor = bitmap.GetPixel(x, y);
        //                        //Console.WriteLine(currentPixelColor.Name);
        //                        if (desiredPixelColor.Name == currentPixelColor.Name ||
        //                            desiredPixelColor2.Name == currentPixelColor.Name ||
        //                            (currentPixelColor.R <= 9 && currentPixelColor.G >= 215 && currentPixelColor.B <= 87))
        //                        {
        //                            //MessageBox.Show($"Found pixel at {x}, {y}");
        //                            coordinates.X = x + 78;
        //                            coordinates.Y = y;
        //                            return coordinates;
        //                        }
        //                    }
        //                }
        //            }
        //            catch (Exception ex)
        //            { }

        //            return coordinates;
        //        }

        //public static Coordinates SearchClipButton(Coordinates prevCoordinates)
        //{
        //    //Reference of previous coordinates
        //    Coordinates coordinates = new Coordinates();

        //    try
        //    {
        //        var initialX = 55;
        //        var initialY = 0;
        //        coordinates.X = prevCoordinates.X + initialX;
        //        coordinates.Y = prevCoordinates.Y + initialY;

        //        //Find bottom field coordinates
        //        coordinates = SearchColor(BottomColor, 1, SystemInformation.VirtualScreen.Height, coordinates.X, coordinates.Y);
        //        initialY = coordinates.Y;

        //        //Return clip field coordinates according to last position
        //        coordinates = SearchColor(GreyColor1, 70, 70, coordinates.X + 65, coordinates.Y, true);
        //    }
        //    catch (Exception ex)
        //    { }

        //    return coordinates;
        //}
        //public static Tuple<int, int> SearchClipButton(Tuple<int, int> prevTuple)
        //{
        //    //Reference of previous coordinates
        //    Tuple<int, int> coordinates = new Tuple<int, int>(-1, -1);

        //    try
        //    {
        //        var initialX = 55;
        //        var initialY = 0;
        //        coordinates = new Tuple<int, int>(prevTuple.Item1 + initialX, prevTuple.Item2 + initialY);

        //        //Find bottom field coordinates
        //        coordinates = SearchColor(BottomColor, 1, SystemInformation.VirtualScreen.Height, coordinates.Item1, coordinates.Item2);
        //        initialY = coordinates.Item2;

        //        //Return clip field coordinates according to last position
        //        coordinates = SearchColor(GreyColor1, 70, 70, coordinates.Item1 + 65, coordinates.Item2, true);
        //    }
        //    catch (Exception ex)
        //    { }

        //    return coordinates;
        //}

        public static int GetEquivalent(int val)
        {
            return decimal.ToInt32(val - ((val) * increaseRate));
        }

        public static void Click(int x, int y, int times = 1)
        {
            SetCursorPos(GetEquivalent(x), GetEquivalent(y));

            for (int i = 0; i < times; i++)
            {
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
            }
        }

        public static void RightClick(int x, int y, int times = 1)
        {
            SetCursorPos(GetEquivalent(x), GetEquivalent(y));

            for (int i = 0; i < times; i++)
            {
                mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
                mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
            }
        }

        public static void SendCtrlC()
        {
            System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
            SendKeys.SendWait("^(c)");
            System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
        }

        public static void SendCtrlV()
        {
            System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
            SendKeys.SendWait("^(v)");
            System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
        }

        public static void SendCtrlA()
        {
            System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
            SendKeys.SendWait("^(a)");
            System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
        }

        public static void SendEnter()
        {
            System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
            SendKeys.SendWait("{ENTER}");
            System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1));
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
