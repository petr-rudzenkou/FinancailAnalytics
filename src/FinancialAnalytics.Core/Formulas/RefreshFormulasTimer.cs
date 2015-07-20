using System;
using System.Linq;
using FinancialAnalytics.Utils.Options;

namespace FinancialAnalytics.Core.Formulas
{
    public class RefreshFormulasTimer : IRefreshFormulasTimer
    {
        private System.Timers.Timer _autoUpdateTimer;
        private int _defaultInterval = 60000;

        public RefreshFormulasTimer()
        {
            InitTimer();
        }
        public int UpdateInterval 
        {
            get { return (int)_autoUpdateTimer.Interval; }
            set { _autoUpdateTimer.Interval = value; }
        }

        public event EventHandler AutoUpdate;

        private void OnAutoUpdate(object sender, EventArgs e)
        {
            var temp = AutoUpdate;
            if (temp != null)
            {
                temp(sender, e);
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

        public bool Enabled
        {
            get { return _autoUpdateTimer.Enabled; }
            set { _autoUpdateTimer.Enabled = value; }
        }

        private void InitTimer()
        {
            _autoUpdateTimer = new System.Timers.Timer();
            _autoUpdateTimer.Elapsed += OnAutoUpdate;
            _autoUpdateTimer.Interval = _defaultInterval;
            try
            {
                var refreshFrequencyOption = OptionsCacheProvider.Options.FirstOrDefault(x => x.GetType() == typeof (RefreshFrequencyOption)) as RefreshFrequencyOption;
                if (refreshFrequencyOption != null)
                {
                    if (!refreshFrequencyOption.IsSelected)
                        return;

                    if (refreshFrequencyOption.RefreshFrequency.HasValue)
                    {
                        int value;
                        switch (refreshFrequencyOption.RefreshFrequencyMeasure)
                        {
                            case RefreshFrequencyMeasure.Sec:
                                value = refreshFrequencyOption.RefreshFrequency.Value*1000;
                                break;
                            case RefreshFrequencyMeasure.Min:
                                value = refreshFrequencyOption.RefreshFrequency.Value*1000*60;
                                break;
                            case RefreshFrequencyMeasure.Hours:
                                value = refreshFrequencyOption.RefreshFrequency.Value*1000*60*60;
                                break;
                            case RefreshFrequencyMeasure.Days:
                                value = refreshFrequencyOption.RefreshFrequency.Value*1000*60*60*24;
                                break;
                            default:
                                value = 60000;
                                break;
                        }
                        _autoUpdateTimer.Interval = value;
                    }
                }
            }
            catch
            {
            }
            finally
            {
                _autoUpdateTimer.Start();
            }
        }

        public void ResetInterval()
        {
            _autoUpdateTimer.Interval = _defaultInterval;
        }
    }
}
