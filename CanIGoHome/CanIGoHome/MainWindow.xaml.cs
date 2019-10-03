using System;
using System.Configuration;
using System.Drawing;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;
using Application = System.Windows.Application;
using MessageBox = System.Windows.Forms.MessageBox;

namespace CanIGoHome
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private NotifyIcon _notifyIcon = null;
        private IWebEngine _webEngine;
        public IWebEngine WebEngine
        {
            get {
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



        public MainWindow()
        {
            InitializeComponent();
            InitNotifyIcon();
            InitUpdateTimer();
            CalculateDailyTime();
            FetchAndAnalyzeData();
        }

        private void InitUpdateTimer()
        {
            // Create a timer with a five second interval.
            fetchTimer = new System.Timers.Timer();
            // Hook up the Elapsed event for the timer. 
            fetchTimer.Elapsed += (s, e) => { _notifyIcon.Text = GetLeaveTimeMsg(); };
            fetchTimer.AutoReset = true;
            fetchTimer.Enabled = false;
        }
        private void InitNotifyIcon()
        {
            _notifyIcon = new NotifyIcon();
            _notifyIcon.Click += new EventHandler(NotifyIcon_Click);
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("CanIGoHome.Resources.RedOnBlackClock.ico"))
            {
                _notifyIcon.Icon = new Icon(stream);
            }
            _notifyIcon.Visible = true;
            _notifyIcon.Text = "Getting entrance time from Hilan web site";

            BuildContextMenu();
        }

        private void BuildContextMenu()
        {
            var ctxMenu = new ContextMenu();
            ctxMenu.MenuItems.Add("Can I Go Home (v0.1)", (s, e) => FetchAndAnalyzeData());
            ctxMenu.MenuItems.Add("--------------------");
            ctxMenu.MenuItems.Add("Fetch Data", (s, e) => FetchAndAnalyzeData());
            ctxMenu.MenuItems.Add("Exit", (s, e) => Application.Current.Shutdown());

            _notifyIcon.ContextMenu = ctxMenu;
        }

        private void CalculateDailyTime()
        {
            if (DateTime.Now.DayOfWeek == DayOfWeek.Thursday)
            {
                _dailyTime = new TimeSpan(7, 30, 0);
            }
            else
            {
                _dailyTime = new TimeSpan(9, 15, 0);
            }
        }

        private void FetchAndAnalyzeData()
        {
            _notifyIcon.Text = "Getting entrance time from Hilan web site";
            _dataFetched = false;
            fetchTimer.Enabled = false;

            var results = FetchData();
            if (string.IsNullOrWhiteSpace(results))
            {
                _notifyIcon.Text = "Failed to fetch information";
                return;
            }

            if (results.Equals("--:--"))
            {
                _notifyIcon.Text = "Entrance time not found";
                return;
            }

            _entranceTime = DateTime.Parse(results);
            _dataFetched = true;
            fetchTimer.Enabled = true;

            _notifyIcon.Text = GetLeaveTimeMsg();
        }

        private string FetchData()
        {
            string user = ConfigurationManager.AppSettings["user"];
            string pw = ConfigurationManager.AppSettings["pw"];
            try
            {
                var result = WebEngine.Search(user, pw);
                return result;
            }
            catch
            {
                return "";
            }
        }

        private void NotifyIcon_Click(object sender, EventArgs e)
        {
            var mouseEvent = (MouseEventArgs)e;
            if (mouseEvent.Button == MouseButtons.Left)
            {
                var info = "No data";
                if (_dataFetched)
                {
                    info = $"You entered at {_entranceTime.ToShortTimeString()}{Environment.NewLine}{Environment.NewLine}" + GetLeaveTimeMsg();
                }
                MessageBox.Show(info, "Can I Go Home (v0.1)");
            }
        }

        private string GetLeaveTimeMsg()
        {
            string leaveTime = (_entranceTime + _dailyTime).ToShortTimeString();
            string leaveHours = (_dailyTime - (DateTime.Now - _entranceTime)).ToString(@"hh\:mm");
            return $"You can leave in {leaveHours} hours (at {leaveTime})";
        }
    }
}
