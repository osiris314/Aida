using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Data.SqlClient;
using System.IO.Ports;
using System.Xml;
using System.Diagnostics;
using System.IO;

namespace test_00
{

    public partial class Form1 : Form
    {
        SpeechSynthesizer s = new SpeechSynthesizer();

        Boolean wake = true;
        public Boolean search = false;
        String temp;
        String condition;

        Choices list = new Choices();
        public Form1()


        {
            SpeechRecognitionEngine rec = new SpeechRecognitionEngine();
            list.Add(File.ReadAllLines(@"commands.txt"));

            Grammar gr = new Grammar(new GrammarBuilder(list));
            try
            {
                rec.RequestRecognizerUpdate();
                rec.LoadGrammar(gr);
                rec.SpeechRecognized += rec_speachrecognized;
                rec.SetInputToDefaultAudioDevice();
                rec.RecognizeAsync(RecognizeMode.Multiple);
            }

            catch
            {
                return;
            }

            s.SelectVoiceByHints(VoiceGender.Female);
            TimeSpan morning1 = new TimeSpan(5, 0, 0); //5 o'clock
            TimeSpan morning2 = new TimeSpan(11, 59, 59); //11:59:59 o'clock
            TimeSpan afternoon1 = new TimeSpan(12, 0, 0); //12 o'clock
            TimeSpan afternoon2 = new TimeSpan(16, 59, 59); //16:59:59 o'clock
            TimeSpan evening1 = new TimeSpan(17, 0, 0); //17 o'clock
            TimeSpan evening2 = new TimeSpan(20, 59, 59); //20:59:59 o'clock
            TimeSpan night1 = new TimeSpan(21, 0, 0); //21 o'clock
            TimeSpan night2 = new TimeSpan(4, 59, 59); //4:59:59 o'clock
            TimeSpan now = DateTime.Now.TimeOfDay;//present time
                if ((now > morning1) && (now < morning2))
                {
                    say("Good morning sir!");
                }

                else if ((now > afternoon1) && (now < afternoon2))
                {
                    say("Good afternoon sir!");
                }

                else if ((now > evening1) && (now < evening2))
                {
                    say("Good evening Sir, how are you?");
                }

                else
                {
                    say("How was your day sir?");
                }
                InitializeComponent();
            //serialPort1.Open();
         say("The sky in syros is " + GetWeather("cond") + "at " + GetWeather("temp") + "degrees");


        }


        private static void Restart()
        {
            ProcessStartInfo proc = new ProcessStartInfo();
            proc.WindowStyle = ProcessWindowStyle.Hidden;
            proc.FileName = "cmd";
            proc.Arguments = "/C shutdown -f -r";
            Process.Start(proc);
        }

        public static void killprog(string s)
        {
            System.Diagnostics.Process[] procs = null;

            try
            {
                procs = Process.GetProcessesByName(s);
                Process prog = procs[0];

                if (!prog.HasExited)
                {
                    prog.Kill();
                }
            }
            finally
            {
                if(procs != null)
                {
                    foreach (Process p in procs)
                    {
                        p.Dispose();
                    }
                }
            }
            
        }

        public String GetWeather(String input)
        {
            String query = String.Format("https://query.yahooapis.com/v1/public/yql?q=select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='hermoupolis, Aegean') and u='c' & format=xml&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys");
            XmlDocument wData = new XmlDocument();
            wData.Load(query);
        
            XmlNamespaceManager manager = new XmlNamespaceManager(wData.NameTable);
            manager.AddNamespace("yweather", "http://xml.weather.yahoo.com/ns/rss/1.0");

            XmlNode channel = wData.SelectSingleNode("query").SelectSingleNode("results").SelectSingleNode("channel");
            XmlNodeList nodes = wData.SelectNodes("query/results/channel");
            try
            {
                temp = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["temp"].Value;
                condition = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["text"].Value;
                
                if (input == "temp")
                {
                    return temp;
                }
                if (input == "cond")
                {
                    return condition;
                }
            }
            catch
            {
                return "Error Reciving data";
            }
            return "error";
        }

        public void say(String h)
        {
            s.Speak(h);
        }

        //friendly social profile
        String[] hello = new String[3] {"Yes sir?", "How can i help", "how can i be at assistance"};

        public String hello_action()
        {
            Random r = new Random();
            return hello[r.Next(3)];
        }




        private void rec_speachrecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string r = e.Result.Text;

            //variables here
            if (r == "Wake up aida") wake = true;
            if (r == "Aida dont listen for a while") wake = false;

            if (search)
            {
                Process.Start("https://www.google.gr/?gfe_rd=cr&ei=19oLWYa3FKWg8wfojKiIBg&gws_rd=ssl#q=" + r);
                search = false;
            }

            if (wake == true && search == false)
            {

                if (r == "search for")
                {
                    search = true;
                }

                //social profile
                if (r == "Aida")
                {
                    say(hello_action());
                }

                if (r == "Thanks" || r == "Thank you")
                {
                    say("your welcome sir!");
                }

                if (r == "What time is it")
                {
                    say(DateTime.Now.ToString("h:mm tt"));
                    if (r == "Come again")
                    {
                        say(DateTime.Now.ToString("h:mm tt"));
                        say("sir");
                    }
                }

                if (r == "Minimize")
                {
                    this.WindowState = FormWindowState.Minimized;
                }

                if (r == "Maximize")
                {
                    this.WindowState = FormWindowState.Normal;
                }

                if (r == "What is today")
                {
                    say(DateTime.Now.ToString("M/d/yyyy"));
                }

                if (r == "What is the weather like")
                {
                    say("The sky in syros is " + GetWeather("cond") + "at " + GetWeather("temp") + "degrees");
                }

                if (r == "What is the temperature")
                {
                    say("it is " + GetWeather("temp") + "degrees");
                }


                //arduino interaction
                if (r == "Open light")
                {
                    //serialPort1.Write("A");
                }

                if (r == "Close light")
                {
                    //serialPort1.Write("B");
                }

                //helper
                if (r == "Create new command")
                {
                    Clipboard.SetText("if (r == )" + Environment.NewLine);
                    SendKeys.Send("^{v}");
                    Clipboard.SetText("{" + Environment.NewLine + Environment.NewLine);
                    SendKeys.Send("^{v}");
                    Clipboard.SetText("}");
                    SendKeys.Send("^{v}");
                }

                if (r == "Restart the pc aida")
                {
                    Restart();
                }


                //program termination
                TimeSpan morning1 = new TimeSpan(5, 0, 0); //5 o'clock
                TimeSpan morning2 = new TimeSpan(11, 59, 59); //11:59:59 o'clock
                TimeSpan afternoon1 = new TimeSpan(12, 0, 0); //12 o'clock
                TimeSpan afternoon2 = new TimeSpan(16, 59, 59); //16:59:59 o'clock
                TimeSpan evening1 = new TimeSpan(17, 0, 0); //17 o'clock
                TimeSpan evening2 = new TimeSpan(20, 59, 59); //20:59:59 o'clock
                TimeSpan night1 = new TimeSpan(21, 0, 0); //21 o'clock
                TimeSpan night2 = new TimeSpan(4, 59, 59); //4:59:59 o'clock
                TimeSpan now = DateTime.Now.TimeOfDay;//present time
                if (r == "Close application")
                {
                    if ((now > morning1) && (now < morning2))
                    {
                        say("Have a nice morning sir!");
                        Environment.Exit(0);
                    }

                    else if ((now > afternoon1) && (now < afternoon2))
                    {
                        say("Have a nice afternoon sir!");
                        Environment.Exit(0);
                    }

                    else if ((now > evening1) && (now < evening2))
                    {
                        say("Have a nice evening sir!");
                        Environment.Exit(0);
                    }

                    else
                    {
                        say("Have a nice night sir!");
                        Environment.Exit(0);
                    }

                }


                //Websites
                if (r == "Open google")
                {
                    System.Diagnostics.Process.Start("https://www.google.com");
                }

                //multi media
                if (r == "Open spotify" || r == "Switch to spotify")
                {
                    Process.Start(@"C:\Users\baggs\AppData\Roaming\Spotify\Spotify.exe");
                }

                if (r == "Open chrome" || r == "Switch to chrome")
                {
                    Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe");
                }

                if (r == "Open visual studio")
                {
                    Process.Start(@"C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\Common7\IDE\devenv.exe");
                }

                if (r == "Close visual studio")
                {
                    killprog("devenv");
                    killprog("devenv");
                    killprog("devenv");
                    killprog("devenv");
                    killprog("devenv");
                }
                if (r == "Close spotify")
                {
                    killprog("Spotify");
                    killprog("Spotify");
                    killprog("Spotify");
                    killprog("Spotify");
                    killprog("Spotify");
                }

                if (r == "Close chrome")
                {
                    killprog("Chrome");
                    killprog("Chrome");
                    killprog("Chrome");
                    killprog("Chrome");
                }

                if (r == "Play" || r == "Stop")
                {
                    SendKeys.Send(" ");
                }

                if (r == "Next" || r == "Next song")
                {
                    SendKeys.Send("^{RIGHT}");
                }

                if (r == "Last" || r == "Last song")
                {
                    SendKeys.Send("^{LEFT}");
                }

                if (r == "Forward")
                {
                    SendKeys.Send("+{RIGHT}");
                    SendKeys.Send("+{RIGHT}");
                    SendKeys.Send("+{RIGHT}");
                }


            }




        }
    }
}
    
