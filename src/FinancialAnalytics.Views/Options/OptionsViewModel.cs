using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Caliburn.Micro;
using FinancialAnalytics.Core.Formulas;
using FinancialAnalytics.Utils.Options;
using FinancialAnalytics.Views.Options.Interfaces;
using OptionsCacheProvider = FinancialAnalytics.Utils.Options.OptionsCacheProvider;

namespace FinancialAnalytics.Views.Options
{
    public class OptionsViewModel : Screen, IOptionsViewModel
    {
        private readonly BindableCollection<OptionBase> _options = new BindableCollection<OptionBase>();
        private readonly List<OptionBase> _defaultOptions = new List<OptionBase>();
        private readonly IRefreshFormulasTimer _refreshFormulasTimer;
        private readonly IDailyRefreshTimer _dailyRefreshTimer;
        public OptionsViewModel(IRefreshFormulasTimer refreshFormulasTimer, IDailyRefreshTimer dailyRefreshTimer)
        {
            _refreshFormulasTimer = refreshFormulasTimer;
            _dailyRefreshTimer = dailyRefreshTimer;
            DisplayName = Resources.ViewsResources.Options_WindowTitle;
        }

        public void OK()
        {
            try
            {
                foreach (var option in _options)
                {
                    var refreshFrequencyOption = option as RefreshFrequencyOption;
                    if (refreshFrequencyOption != null)
                    {
                        if (option.IsSelected)
                        {
                            if (refreshFrequencyOption.RefreshFrequency.HasValue)
                            {
                                int value;
                                switch (refreshFrequencyOption.RefreshFrequencyMeasure)
                                {
                                    case RefreshFrequencyMeasure.Sec:
                                        value = refreshFrequencyOption.RefreshFrequency.Value * 1000;
                                        break;
                                    case RefreshFrequencyMeasure.Min:
                                        value = refreshFrequencyOption.RefreshFrequency.Value * 1000 * 60;
                                        break;
                                    case RefreshFrequencyMeasure.Hours:
                                        value = refreshFrequencyOption.RefreshFrequency.Value * 1000 * 60 * 60;
                                        break;
                                    case RefreshFrequencyMeasure.Days:
                                        value = refreshFrequencyOption.RefreshFrequency.Value * 1000 * 60 * 60 * 24;
                                        break;
                                    default:
                                        value = 60000;
                                        break;
                                }

                                _refreshFormulasTimer.UpdateInterval = value;
                                _refreshFormulasTimer.Start();

                                OptionsCacheProvider.Set(refreshFrequencyOption);
                            }
                        }
                        else
                        {
                            _refreshFormulasTimer.ResetInterval();
                            OptionsCacheProvider.UnSet(refreshFrequencyOption);
                        }
                    }

                    var dailyRefreshTimeOption = option as DailyRefreshTimeOption;
                    if (dailyRefreshTimeOption != null)
                    {
                        if (option.IsSelected)
                        {
                            string[] timeStr = dailyRefreshTimeOption.DailyRefreshTime.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries).ToArray();
                            int hour = int.Parse(timeStr[0]);
                            int min = int.Parse(timeStr[1]);
                            int sec = int.Parse(timeStr[2]);

                            if ((hour > 23 || hour < 0) || (min > 60 || min < 0) || (sec > 60 || sec < 0))
                                throw new Exception();

                            _dailyRefreshTimer.Hour = hour;
                            _dailyRefreshTimer.Min = min;
                            _dailyRefreshTimer.Sec = sec;

                            _dailyRefreshTimer.Start();

                            OptionsCacheProvider.Set(dailyRefreshTimeOption);
                        }
                        else
                        {
                            _dailyRefreshTimer.Stop();
                            OptionsCacheProvider.UnSet(dailyRefreshTimeOption);
                        }
                    }
                }
                TryClose();
            }
            catch
            {
                MessageBox.Show("Specify valid parameters");
            }
        }

        protected override void OnViewAttached(object view, object context)
        {
            base.OnViewAttached(view, context);
            CreateDefaultOptions();
            CreateOptions();
        }

        public void Cancel()
        {
            TryClose();
        }

        public BindableCollection<OptionBase> Options
        {
            get { return _options; }
        }

        private void CreateOptions()
        {
            _options.Clear();
            OptionsCacheProvider.Initialize();

            foreach (var option in OptionsCacheProvider.Options)
            {
                if (!_options.Any(x => x.Name.Equals(option.Name)))
                {
                    _options.Add(option);
                }
            }

            foreach (var defaultOption in _defaultOptions)
            {
                if (!_options.Any(x => x.Name.Equals(defaultOption.Name)))
                {
                    _options.Add(defaultOption);
                }
            }
        }

        private void CreateDefaultOptions()
        {
            _defaultOptions.Clear();
            // Create default options
            _defaultOptions.Add(new RefreshFrequencyOption()
            {
                Name = OptionsConstants.RefreshFrequency,
                DisplayName = "Refresh Frequency: ",
                RefreshFrequencyMeasure = RefreshFrequencyMeasure.Sec,
                IsSelected = false
            });
            _defaultOptions.Add(new DailyRefreshTimeOption()
            {
                Name = OptionsConstants.DailyRefreshTime,
                DisplayName = "Daily Refresh Time: ",
                IsSelected = false
            });
        }
    }
}
