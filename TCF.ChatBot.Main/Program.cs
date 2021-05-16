using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TCF.ChatBot.Main.Utils;

namespace TCF.ChatBot.Main
{
    public class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
            //Find unread messages
            Exec();

            //List<bool> iHash0 = Utilities.GetHash(new Bitmap(@"C:\Users\thiag\source\repos\TCF.ChatBot\TCF.ChatBot.Main\Utils\imgGreenCircle.png"));
            //List<bool> iHash1 = Utilities.GetHash(new Bitmap(@"C:\Users\thiag\source\repos\TCF.ChatBot\TCF.ChatBot.Main\Utils\imgGreenCircle1.png"));
            //List<bool> iHash2 = Utilities.GetHash(new Bitmap(@"C:\Users\thiag\source\repos\TCF.ChatBot\TCF.ChatBot.Main\Utils\imgGreenCircle2.png"));
            //List<bool> iHash3 = Utilities.GetHash(new Bitmap(@"C:\Users\thiag\source\repos\TCF.ChatBot\TCF.ChatBot.Main\Utils\imgGreenCircle3.png"));
            //List<bool> iHash4 = Utilities.GetHash(new Bitmap(@"C:\Users\thiag\source\repos\TCF.ChatBot\TCF.ChatBot.Main\Utils\imgGreenCircle4.png"));
            //List<bool> iHash5 = Utilities.GetHash(new Bitmap(@"C:\Users\thiag\source\repos\TCF.ChatBot\TCF.ChatBot.Main\Utils\imgGreenCircle5.png"));
            //List<bool> iHash6 = Utilities.GetHash(new Bitmap(@"C:\Users\thiag\source\repos\TCF.ChatBot\TCF.ChatBot.Main\Utils\imgGreenCircle6.png"));
            //List<bool> iHash7 = Utilities.GetHash(new Bitmap(@"C:\Users\thiag\source\repos\TCF.ChatBot\TCF.ChatBot.Main\Utils\imgGreenCircle7.png"));
            //List<bool> iHash8 = Utilities.GetHash(new Bitmap(@"C:\Users\thiag\source\repos\TCF.ChatBot\TCF.ChatBot.Main\Utils\imgGreenCircle8.png"));
            //List<bool> iHash9 = Utilities.GetHash(new Bitmap(@"C:\Users\thiag\source\repos\TCF.ChatBot\TCF.ChatBot.Main\Utils\imgGreenCircle9.png"));

            ////Determine the number of equal pixel (x of 256)
            //int equalElements1 = iHash0.Zip(iHash1, (i, j) => i == j).Count(eq => eq);
            //int equalElements2 = iHash0.Zip(iHash2, (i, j) => i == j).Count(eq => eq);
            //int equalElements3 = iHash0.Zip(iHash3, (i, j) => i == j).Count(eq => eq);
            //int equalElements4 = iHash0.Zip(iHash4, (i, j) => i == j).Count(eq => eq);
            //int equalElements5 = iHash0.Zip(iHash5, (i, j) => i == j).Count(eq => eq);
            //int equalElements6 = iHash0.Zip(iHash6, (i, j) => i == j).Count(eq => eq);
            //int equalElements7 = iHash0.Zip(iHash7, (i, j) => i == j).Count(eq => eq);
            //int equalElements8 = iHash0.Zip(iHash8, (i, j) => i == j).Count(eq => eq);
            //int equalElements9 = iHash0.Zip(iHash9, (i, j) => i == j).Count(eq => eq);
        }

        private static void Exec()
        {
            var dir = Utilities.AssemblyDirectory;
            var greenCircle = Utilities.FindExactImage(new Bitmap($@"{dir}\Utils\imgNewMessage.png"));

            if (greenCircle.Item1 != -1 && greenCircle.Item2 != -1)
            {
                //Clicks on the unread message
                Utilities.Click(greenCircle.Item1, greenCircle.Item2);

                //Finds paperclip to make tripple click to select the message
                var paperClip = Utilities.FindExactImage(new Bitmap($@"{dir}\Utils\imgPaperClip.png"));
                Utilities.Click((paperClip.Item1 + 30), (paperClip.Item2 - 85), 3);

                //Right click to find "Copy" option
                //Utilities.RightClick((paperClip.Item1 + 30), (paperClip.Item2 - 85));

                Utilities.SendCtrlC();
                var text = Clipboard.GetData(DataFormats.Text)?.ToString();

                //Set focus on type text field
                Utilities.Click((paperClip.Item1 + 70), (paperClip.Item2));
                Utilities.SendCtrlA();

                //Add rule for message
                switch (text?.ToUpper())
                {
                    case "OI":
                        Utilities.SendText("Opa, tudo certo?");
                        break;
                    default:
                        Utilities.SendText("Olá! Esta mensagem esta sendo digitada por uma automação. Não responder");
                        break;
                }

                Utilities.SendEnter();
            }
        }


        public static List<bool> GetHash(Bitmap bmpSource)
        {
            List<bool> lResult = new List<bool>();
            //Create new image with 16x16 pixel
            Bitmap bmpMin = new Bitmap(bmpSource, new Size(16, 16));
            for (int j = 0; j < bmpMin.Height; j++)
            {
                for (int i = 0; i < bmpMin.Width; i++)
                {
                    var pixel = bmpMin.GetPixel(i, j);
                    var brightness = pixel.GetBrightness();
                    //Reduce Colors to True / False                
                    lResult.Add(brightness < 0.5f);
                }
            }
            return lResult;
        }
    }
}
