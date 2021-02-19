using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace TeamsAutoJoiner
{

    //Todo:
    //Chrome cleanup
    class Teams
    {
        static String SettingsFile = "config.json";

        private struct MeetingEntry
        {
            public TimeSpan start;
            public TimeSpan end;
            public Uri url;
            public String label;
            public String comment;
            public MeetingEntry(MeetingEntry entry)
            {
                start = entry.start;
                end = entry.end;
                url = entry.url;
                label = entry.label;
                comment = entry.comment;
            }
        };

        Queue<MeetingEntry> Meetings;
        MeetingEntry ActiveMeeting;

        class Settings
        {
            public String SrcFile = "schedule.txt";
            public String username = "";
            public String password = "";
            public bool headless = true;
            public bool mute = true;

            bool validate()
            {
                return username != "" && password != "";
            }
        };
        Settings options = new Settings();

        IWebDriver driver = null;
        ChromeDriverService service = null;
        ChromeOptions opt = null;

        private void InitChrome()
        {
            service = ChromeDriverService.CreateDefaultService();
            opt = new ChromeOptions();
            
            opt.AddArguments(new string[] {
                "--no-sandbox",
                "--disable-password-manager",
                "--disable-speech-api",
                "--disable-default-apps",
                "--disable-infobars",
                "--disable-extensions",
                "--use-fake-ui-for-media-stream"
            });

            string pref = "2";
            //Headless by default muted
            if (options.headless) { opt.AddArguments(new string[] { "--headless", "--disable-gpu" }); }
            else
            { //Muting audio and mic
                if (options.mute)
                {
                    opt.AddArgument("--mute-audio");
                }
                else pref = "1";
            }
            opt.AddLocalStatePreference("profile.default_content_setting_values.media_stream_mic", pref);
            //Camera always off
            opt.AddLocalStatePreference("profile.default_content_setting_values.media_stream_camera", "2");
            //Notifications always off
            opt.AddLocalStatePreference("profile.default_content_setting_values.notifications", "2");


            //Log stuff
            Directory.CreateDirectory("Logs");
            File.CreateText("chromedriver.log").Close();

            service.EnableVerboseLogging = false;
            service.LogPath = Path.Combine("Logs", "chromedriver.log");

            //Disable messages in console
            service.HideCommandPromptWindow = true;

            driver = new ChromeDriver(service, opt);

            //Set 20 sec timeout
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
        }
        private void ParseArgs(String []args)
        {
            if (args.Length == 0) return;
            if (args.Length == 1) { options.SrcFile = args[0]; return; }
            for (uint i = 0; i < args.Length-1; i++)
            {
                if (args[i].StartsWith("-") && args[i].Length > 2)
                {
                    switch(args[i][1])
                    {
                        case 'F':
                        case 'f':
                            options.SrcFile = args[++i];
                            break;
                        case 'C':
                        case 'c':
                            SettingsFile = args[++i];
                            break;
                        case 'u':
                        case 'U':
                            options.username = args[++i];
                            break;
                        case 'p':
                        case 'P':
                            options.password = args[++i];
                            break;
                        case 's':
                        case 'S':
                            SaveConfig();
                            break;
                    }
                }
            }
        }
        private void LoadSchududle()
        {
            if (!File.Exists(options.SrcFile)) throw new FileNotFoundException("File " + options.SrcFile + " dont exists.");

            StreamReader fileStream = File.OpenText(options.SrcFile);
            Meetings = new Queue<MeetingEntry>();
            while(!fileStream.EndOfStream)
            {
                String line = fileStream.ReadLine();
                MeetingEntry me = new MeetingEntry();
                int idxb = line.LastIndexOf('(');
                int idxe = line.LastIndexOf(')');
                int diff = idxe - idxb;
                me.label = line.Substring(0, idxb);
                String[] time = line.Substring(idxb + 1, diff-1).Split('-');
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write(me.label + " ");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(time[0] + " " + time[1]);
                me.start = TimeSpan.ParseExact(time[0], "h\\:mm", System.Globalization.CultureInfo.InvariantCulture);
                me.end = TimeSpan.ParseExact(time[1], "h\\:mm", System.Globalization.CultureInfo.InvariantCulture);
                String url = fileStream.ReadLine();
                if (url.StartsWith("https://")) me.url = new Uri(url);
                else me.comment = url;

                Meetings.Enqueue(me);
            }
            fileStream.Close();
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine("Shedule loaded.");
            Console.ResetColor();
        }
        public void Work(String []args)
        {
            this.ParseArgs(args);
            this.LoadConfig();
            this.LoadSchududle();

            this.InitChrome();

            this.ActivateMeeting();
        }
        private void SaveConfig()
        {
            try
            {
                using (StreamWriter f = new StreamWriter(SettingsFile))
                {
                    f.Write(JsonConvert.SerializeObject(options));
                }
            }
            catch (Exception ex) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("Coudnt save config"); throw new ArgumentException(); }
            Console.ResetColor();
        }
        private void LoadConfig()
        {
            try
            {
                using (StreamReader f = new StreamReader(SettingsFile))
                {
                    options = JsonConvert.DeserializeObject<Settings>(f.ReadToEnd());
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine("Config Loaded.");
                }
            } catch(Exception ex) {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Coudnt load config");
                throw new ArgumentException();
            }
            finally
            {
                SaveConfig();
                Console.ResetColor();
            }
        }
        private void Login()
        {
            Thread.Sleep(1000);
            driver.FindElement(By.Id("i0116")).SendKeys(options.username + Keys.Enter);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("i0118")).SendKeys(options.password + Keys.Enter);
            driver.FindElement(By.Id("idSIButton9")).Click();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Login successful.");
        }
        private void ActivateMeeting()
        {
            while (Meetings.Count != 0)
            {
                if (driver.Url != "data:,") driver.Navigate().GoToUrl("data:,");

                ActiveMeeting = Meetings.Dequeue();
                if (ActiveMeeting.url == null) continue;

                //Time managment
                if (ActiveMeeting.url != null && ActiveMeeting.start.CompareTo(DateTime.Now.TimeOfDay) == 1)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Waiting to start " + ActiveMeeting.label + " in " + ActiveMeeting.start + ".");
                    Task.Delay(ActiveMeeting.start - DateTime.Now.TimeOfDay).Wait();
                }
                else if (!(DateTime.Now.TimeOfDay - ActiveMeeting.start < ActiveMeeting.end - ActiveMeeting.start))
                {
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    Console.WriteLine(ActiveMeeting.label + " gone.");
                    continue;
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Connecting to " + ActiveMeeting.label + ".");

                driver.Navigate().GoToUrl(ActiveMeeting.url);

                driver.FindElements(By.CssSelector("button.btn.primary"))[1].Click();

                //Try login
                try
                {
                    driver.FindElement(By.LinkText("войти")).Click();
                    Login();
                }
                catch(NoSuchElementException ex) { }


                driver.FindElement(By.CssSelector("button.join-btn.ts-btn.inset-border.ts-btn-primary")).Click();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Connected.");


                Task.Delay(ActiveMeeting.end - ActiveMeeting.start).Wait();
                Console.WriteLine(ActiveMeeting.label + " Ended.");
            }
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine("No meetings.");
            SaveConfig();

            Console.ResetColor();
        }
        public void Close()
        {
            Console.WriteLine("Exiting.");

            //Fix background browser stays open
            service.HideCommandPromptWindow = false;
            service.Dispose();
            driver.Quit();
        }
    }
}
