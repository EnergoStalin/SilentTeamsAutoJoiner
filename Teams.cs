using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace TeamsAutoJoiner
{
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
            public String SrcFile = "";
            public String logLevel = "3";
            public String username = "";
            public String password = "";

            bool validate()
            {
                int _;
                return int.TryParse(logLevel, out _) && username != "" && password != "" && SrcFile != "";
            }
        };
        Settings options = new Settings();

        IWebDriver driver = null;

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
                        case 'd':
                        case 'D':
                            options.logLevel = "1";
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

            Console.WriteLine("Loading shedudle.");

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
                Console.WriteLine(me.label + " " + time[0] + " " + time[1]);
                me.start = TimeSpan.ParseExact(time[0], "h\\:mm", System.Globalization.CultureInfo.InvariantCulture);
                me.end = TimeSpan.ParseExact(time[1], "h\\:mm", System.Globalization.CultureInfo.InvariantCulture);
                String url = fileStream.ReadLine();
                if (url.StartsWith("https://")) me.url = new Uri(url);
                else me.comment = url;

                Meetings.Enqueue(me);
            }
            fileStream.Close();
        }
        public void Work(String []args)
        {
            this.ParseArgs(args);
            this.LoadConfig();
            this.LoadSchududle();

            ChromeOptions opt = new ChromeOptions();
            opt.AddLocalStatePreference("profile.default_content_setting_values.media_stream_mic", "2");
            opt.AddLocalStatePreference("profile.default_content_setting_values.media_stream_camera", "2");
            opt.AddLocalStatePreference("profile.default_content_setting_values.notifications", "2");
            opt.AddArguments(new string[] { "--log-level=" + options.logLevel, "--no-sandbox", "--headless", "--disable-speech-api", "--mute-audio", "--disable-default-apps", "--disable-infobars", "--disable-extensions", "use-fake-ui-for-media-stream" });

            driver = new ChromeDriver(opt);
            _ = driver.Manage().Timeouts().ImplicitWait;

            this.ActivateMeeting();
        }
        private void SaveConfig()
        {
            try
            {
                using (StreamWriter f = new StreamWriter(SettingsFile))
                {
                    f.Write(JsonConvert.SerializeObject(options));
                    Console.WriteLine("Config saved");
                }
            }
            catch (Exception ex) { Console.WriteLine("Coudnt save config"); throw new ArgumentException(); }
        }
        private void LoadConfig()
        {
            try
            {
                using (StreamReader f = new StreamReader(SettingsFile))
                {
                    options = JsonConvert.DeserializeObject<Settings>(f.ReadToEnd());
                    Console.WriteLine("Config Loaded");
                }
            } catch(Exception ex) { SaveConfig(); Console.WriteLine("Coudnt load config"); throw new ArgumentException(); }
        }
        private void Login(WebDriverWait wt)
        {
            wt.Until(driver => driver.FindElement(By.Id("i0116"))).SendKeys(options.username + Keys.Enter);
            Thread.Sleep(1000);
            wt.Until(driver => driver.FindElement(By.Id("i0118"))).SendKeys(options.password + Keys.Enter);
            Thread.Sleep(1000);
            wt.Until(driver => driver.FindElement(By.Id("idSIButton9"))).Click();

            Console.WriteLine("Login successful");
            Console.WriteLine("Username is {}", options.username);
            Console.WriteLine("Password is {}", options.password);
        }
        private void ActivateMeeting()
        {
            while(Meetings.Count != 0)
            {
                if(driver.Url != "data:,") driver.Navigate().GoToUrl("data:,");

                ActiveMeeting = Meetings.Dequeue();
                if (ActiveMeeting.url == null) continue;

                if (ActiveMeeting.url != null && ActiveMeeting.start.CompareTo(DateTime.Now.TimeOfDay) == 1)
                {
                    Console.WriteLine("Waiting to start meeting in " + ActiveMeeting.start);
                    Task.Delay(ActiveMeeting.start - DateTime.Now.TimeOfDay).Wait();
                }
                else if (!(DateTime.Now.TimeOfDay - ActiveMeeting.start < ActiveMeeting.end - ActiveMeeting.start))
                {
                    Console.WriteLine(ActiveMeeting.label + " is gone");
                    continue;
                }
                Console.WriteLine("Connecting to " + ActiveMeeting.label);

                Task wait = Task.Delay(ActiveMeeting.end - ActiveMeeting.start);
                WebDriverWait wt = new WebDriverWait(driver, TimeSpan.FromSeconds(60));

                driver.Navigate().GoToUrl(ActiveMeeting.url);

                wt.Until(driver => driver.FindElements(By.CssSelector("button.btn.primary")))[1].Click();

                try
                {
                    new WebDriverWait(driver, TimeSpan.FromSeconds(7)).Until(ExpectedConditions.ElementExists(By.LinkText("войти")));
                    driver.FindElement(By.LinkText("войти")).Click();
                    Thread.Sleep(4000);
                    Login(wt);
                }
                catch(Exception ex) {}

                wt.Until(driver => driver.FindElement(By.CssSelector("button.join-btn.ts-btn.inset-border.ts-btn-primary"))).Click();
                Console.WriteLine("Connected");


                wait.Wait();
            }
            Console.WriteLine("No meetings");
            SaveConfig();
        }
        ~Teams()
        {
            driver?.Close();
        }
    }
}
