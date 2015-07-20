using System;
using System.Linq;
using FinancialAnalytics.Utils.Options;

namespace FinancialAnalytics.Core.Formulas
{
    public class DailyRefreshTimer : IDailyRefreshTimer
    {
        private System.Timers.Timer _autoUpdateTimer;

        public DailyRefreshTimer()
        {
            InitTimer();
        }

        public event EventHandler AutoUpdate;

        private void OnElapsed(object sender, EventArgs e)
        {
            if (Hour.HasValue && Min.HasValue && Sec.HasValue)
            {
                if (DateTime.Now.Hour == Hour.Value && DateTime.Now.Minute == Min.Value && DateTime.Now.Second == Sec.Value)
                {
                    var temp = AutoUpdate;
                    if (temp != null)
                    {
                        temp(sender, e);
                    }
                }
            }
        }

        public void Start()
        {
            _autoUpdateTimer.Start();
        }

        public void Stop()
        {
            _autoUpdateTimer.Stop();
        }

        private void InitTimer()
        {
            _autoUpdateTimer = new System.Timers.Timer();
            _autoUpdateTimer.Interval = 1000;
            _autoUpdateTimer.Elapsed += OnElapsed;
           
            try
            {
                var dailyRefreshTimeOption = OptionsCacheProvider.Options.FirstOrDefault(x => x.GetType() == typeof(DailyRefreshTimeOption)) as DailyRefreshTimeOption;
                if (dailyRefreshTimeOption != null)
                {
                    if (!dailyRefreshTimeOption.IsSelected)
                        return;

                    string[] timeStr = dailyRefreshTimeOption.DailyRefreshTime.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries).ToArray();
                    int hour = int.Parse(timeStr[0]);
                    int min = int.Parse(timeStr[1]);
                    int sec = int.Parse(timeStr[2]);

                    Hour = hour;
                    Min = min;
                    Sec = sec;

                    _autoUpdateTimer.Start();
                }
            }
            catch
            {
            }
        }

        public int? Sec { get; set; }

        public int? Min { get; set; }

        public int? Hour { get; set; }


        public bool Enabled
        {
            get { return _autoUpdateTimer.Enabled; }
            set { _autoUpdateTimer.Enabled = value; }
        }
    }
}
