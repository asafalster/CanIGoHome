using System;
using System.Collections.Specialized;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Application = System.Windows.Application;

namespace CanIGoHome
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private readonly string _version = "0.2";
        private NotifyIcon _notifyIcon = null;
        private IWebEngine _webEngine;
        public IWebEngine WebEngine
        {
            get
            {
                if (_webEngine == null)
                {
                    _webEngine = new SeleniumSimpleWebEngine();
                }
                return _webEngine;
            }
        }

        private TimeSpan _dailyTime;
        private DateTime _entranceTime;
        private bool _dataFetched = false;
        private System.Timers.Timer fetchTimer;

        public App()
        {
            logger.Debug("App is starting");
            EncryptAppConfigData();
            InitNotifyIcon();
            InitUpdateTimer();
            GetDailyTime();
            ConfigWebEngine();
            FetchAndAnalyzeData();
            logger.Debug("App init ended");
        }

        private void ConfigWebEngine()
        {
            var appConfiguration = ((NameValueCollection)ConfigurationManager.GetSection("AppConfiguration")).ToDictionary();
            WebEngine.ConfigureEngine(ConfigurationManager.AppSettings.ToDictionary(appConfiguration));
        }

        private void EncryptAppConfigData()
        {
            logger.Trace($"Enter");
            var config = ConfigurationManager.OpenExeConfiguration(Assembly.GetExecutingAssembly().Location);
            
            var section = config.GetSection("appSettings");
            if (!section.SectionInformation.IsProtected)
            {
                // Encrypt the section.
                section.SectionInformation.ProtectSection("DataProtectionConfigurationProvider");
                config.Save();
            }
            logger.Trace($"Exit");
        }

        private void InitUpdateTimer()
        {
            logger.Trace($"Enter");
            // Create a timer with a five second interval.
            fetchTimer = new System.Timers.Timer(60000);
            // Hook up the Elapsed event for the timer. 
            fetchTimer.Elapsed += (s, e) => { _notifyIcon.Text = GetLeaveTimeMsg(); };
            fetchTimer.AutoReset = true;
            fetchTimer.Enabled = false;
            logger.Trace($"Exit");
        }

        private void InitNotifyIcon()
        {
            logger.Trace($"Enter");
            _notifyIcon = new NotifyIcon();
            _notifyIcon.Click += new EventHandler(NotifyIcon_Click);
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("CanIGoHome.Resources.RedOnBlackClock.ico"))
            {
                _notifyIcon.Icon = new Icon(stream);
            }
            _notifyIcon.Visible = true;
            _notifyIcon.Text = "Getting entrance time from Hilan web site";

            BuildContextMenu();
            logger.Trace($"Exit");
        }

        private void BuildContextMenu()
        {
            logger.Trace($"Enter");
            var ctxMenu = new ContextMenu();
            ctxMenu.MenuItems.Add($"Can I Go Home (v{_version})", (s, e) => FetchAndAnalyzeData());
            ctxMenu.MenuItems.Add("--------------------");
            ctxMenu.MenuItems.Add("Fetch Data", (s, e) => FetchAndAnalyzeData());
            ctxMenu.MenuItems.Add("Open log folder", (s, e) => OpenLogFolder());
            ctxMenu.MenuItems.Add("Exit", (s, e) => Current.Shutdown());

            _notifyIcon.ContextMenu = ctxMenu;
            logger.Trace($"Exit");
        }

        private void OpenLogFolder()
        {
            logger.Trace($"Enter");
            var path = $"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}{Path.DirectorySeparatorChar}logs";
            if (Directory.Exists(path))
            {
                Process.Start(path);
            }
            else
            {
                logger.Debug($"Logs folder does not exist - {path}");
            }            
            logger.Trace($"Exit");
        }

        private void GetDailyTime()
        {
            logger.Trace($"Enter");
            var dayOfWeek = (int)DateTime.Now.DayOfWeek + 1;
            var duration = GetSettingsValue<string>($"Day{dayOfWeek}Duration", "AppConfiguration");
            if (!TimeSpan.TryParse(duration, out _dailyTime))
            {
                logger.Debug($"Failed to parse todays duration. Using default");
                if (DateTime.Now.DayOfWeek == DayOfWeek.Thursday)
                {
                    _dailyTime = new TimeSpan(7, 30, 0);
                }
                else
                {
                    _dailyTime = new TimeSpan(9, 15, 0);
                }
            }
            logger.Trace($"Exit");
        }

        private T GetSettingsValue<T>(string valueName,string sectionName)
        {
            logger.Trace($"Enter");
            var configSection = ((NameValueCollection)ConfigurationManager.GetSection(sectionName));
            var ret = (T)Convert.ChangeType(configSection[valueName], typeof(T));
            logger.Trace($"Exit");
            return ret;
        }

        private void FetchAndAnalyzeData()
        {
            logger.Trace($"Enter");
            if (DateTime.Now.DayOfWeek == DayOfWeek.Friday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
            {
                _notifyIcon.Text = "It is weekenend. You are not working today";
                logger.Trace($"Exit");
                return;
            }
            _notifyIcon.Text = "Getting entrance time from Hilan web site";
            _dataFetched = false;
            fetchTimer.Enabled = false;

            var results = FetchData();
            if (string.IsNullOrWhiteSpace(results))
            {
                _notifyIcon.Text = "Failed to fetch information";
                logger.Trace($"Exit");
                return;
            }

            if (results.Equals("--:--"))
            {
                _notifyIcon.Text = "Entrance time not found";
                logger.Trace($"Exit");
                return;
            }

            _entranceTime = DateTime.Parse(results);
            _dataFetched = true;
            fetchTimer.Enabled = true;

            _notifyIcon.Text = GetLeaveTimeMsg();
            logger.Trace($"Exit");
        }

        private string FetchData()
        {
            logger.Trace($"Enter");
            string res = string.Empty;
            try
            {
                res = WebEngine.Search();
            }
            catch(Exception ex)
            {
                logger.Debug(ex,$"WebEngine.Search() throw exception");
                return "";
            }
            logger.Trace($"Exit");
            return res;
        }

        private void NotifyIcon_Click(object sender, EventArgs e)
        {
            logger.Trace($"Enter");
            var mouseEvent = (MouseEventArgs)e;
            if (mouseEvent.Button == MouseButtons.Left)
            {
                var info = "No data";
                if (DateTime.Now.DayOfWeek == DayOfWeek.Friday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
                {
                    info = "It is weekenend. You are not working today";                    
                }
                else if (_dataFetched)
                {
                    info = $"You entered at {_entranceTime.ToShortTimeString()}{Environment.NewLine}{Environment.NewLine}" + GetLeaveTimeMsg();
                }
                MessageBox.Show(info, $"Can I Go Home (v{_version})");
            }
            logger.Trace($"Exit");
        }

        private string GetLeaveTimeMsg()
        {
            logger.Trace($"Enter");
            string leaveTime = (_entranceTime + _dailyTime).ToShortTimeString();
            string leaveHours = (_dailyTime - (DateTime.Now - _entranceTime)).ToString(@"hh\:mm");
            var ret = $"You can leave in {leaveHours} hours (at {leaveTime})";
            logger.Trace($"Exit");
            return ret;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private string GetCurrentMethod()
        {
            var st = new StackTrace();
            var sf = st.GetFrame(1);

            return sf.GetMethod().Name;
        }
    }
}
