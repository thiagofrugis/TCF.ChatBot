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
#if !DEBUG
            //Sets an key to stop the application
            InterceptKeyboard.SetInterceptKeyboard(Keys.Escape);
            //InterceptKeyboard.SetInterceptKeyboard(true, Keys.LControlKey, Keys.F8);
#endif
            while (!InterceptKeyboard.stopExecuting)
            {
                Task.Delay(TimeSpan.FromSeconds(3)).Wait();
                Exec();
            }
        }

        private static void Exec()
        {
            var dir = Utilities.AssemblyDirectory;
            var greenCircle = Utilities.FindExactImage(new Bitmap($@"{dir}\Utils\imgNewMessage.png"));

            if (greenCircle.Item1 != -1 && greenCircle.Item2 != -1)
            {
                //Clicks on the unread message
                Utilities.Click(greenCircle.Item1 - 100, greenCircle.Item2);

                //Finds paperclip as reference and makes tripple click to select the message
                var paperClip = Utilities.FindExactImage(new Bitmap($@"{dir}\Utils\imgPaperClip.png"));
                Utilities.Click((paperClip.Item1 + 20), (paperClip.Item2 - 85), 3);

                //With message selected, send Ctrl+C
                Utilities.SendCtrlC();

                //Reads data copied
                var text = Clipboard.GetData(DataFormats.Text)?.ToString()?.ToUpper();

                //Set focus on type text field
                Utilities.Click((paperClip.Item1 + 70), (paperClip.Item2));
                Utilities.SendCtrlA();

                if (Utilities.ListGreetingsWords.Any(_ => !string.IsNullOrEmpty(text) && text.Contains(_)))
                    Utilities.SendText("Opa, tudo certo?");
                else
                {
                    Utilities.SendText("Esta é uma mensagem automática.");
                    Utilities.SendShiftEnter();
                    Utilities.SendText("Favor não responder");
                }

                Utilities.SendEnter();
            }
        }
    }
}
