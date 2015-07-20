using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Caliburn.Micro;
using FinancialAnalytics.DataFacades.Charts;
using FinancialAnalytics.Views.Charts.CriteriaGroups;
using FinancialAnalytics.Views.Charts.Criterias;
using FinancialAnalytics.Views.Charts.GroupContainers;
using FinancialAnalytics.Views.Charts.Interfaces;

namespace FinancialAnalytics.Views.Charts
{
    public class ChartsInfoViewModel : Screen, IChartsInfoViewModel
    {
        private ChartDownloadSettings _chartSettings;
        private BitmapImage _chart;
        private readonly BindableCollection<CriteriaGroupContainer> _groupContainers = new BindableCollection<CriteriaGroupContainer>();

        public ChartsInfoViewModel(string id)
        {
            DisplayName = id;
            Symbol = id;
            CreateChartCriteriaGroupContainers();
        }

        public string Symbol { get; set; }

        public BitmapImage Chart
        {
            get { return _chart; }
            set
            {
                _chart = value;
                NotifyOfPropertyChange(() => Chart);
            }
        }

        public BindableCollection<CriteriaGroupContainer> GroupContainers
        {
            get { return _groupContainers; }
        }

        private void CreateChartCriteriaGroupContainers()
        {
            GroupContainers.Add(new BasicGroupContainer());
            GroupContainers.Add(new MovingAvgContainer());
            GroupContainers.Add(new IndicatorsContainer());
            GroupContainers.Add(new OverlaysContainer());
        }

        public void GetChart()
        {
            var groups = new List<ChartCriteriaGroup>();
            foreach (var container in GroupContainers)
            {
                groups.AddRange(container.ChartCriteriaGroups);
            }

            var criterias = new List<ChartCriteria>();
            foreach (var group in groups)
            {
                criterias.AddRange(group.ChartCriterias);
            }
            ConstructChartSettings(criterias);
            Chart = new BitmapImage(new Uri(_chartSettings.GetUrl()));
        }

        public void ExecuteGetChart(ActionExecutionContext context)
        {
            var keyArgs = context.EventArgs as KeyEventArgs;
            if (keyArgs == null)
                return;

            if (keyArgs.Key != Key.Enter)
                return;

            var textBox = context.Source as TextBox;
            if (textBox != null)
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                {
                    var groups = new List<ChartCriteriaGroup>();
                    foreach (var container in GroupContainers)
                    {
                        groups.AddRange(container.ChartCriteriaGroups);
                    }

                    var criterias = new List<ChartCriteria>();
                    foreach (var group in groups)
                    {
                        criterias.AddRange(group.ChartCriterias);
                    }
                    var selectedCriterias = criterias.Where(x => x.IsSelected).ToList();
                    var cmpVsCriteria = selectedCriterias.FirstOrDefault(x => x.GetType() == typeof(CompareVsCriteria)) as CompareVsCriteria;
                    if (cmpVsCriteria != null)
                    {
                        cmpVsCriteria.Ids = textBox.Text;
                        GetChart();
                    }
                }
            }
        }

        private void ConstructChartSettings(IEnumerable<ChartCriteria> criterias)
        {
            _chartSettings = new ChartDownloadSettings();
            _chartSettings.ID = Symbol;
            var selectedCriterias = criterias.Where(x => x.IsSelected).ToList();
            foreach (var criteria in selectedCriterias)
            {
                var rangeCriteria = criteria as RangeCriteria;
                if (rangeCriteria != null)
                {
                    _chartSettings.TimeSpan = rangeCriteria.Range;
                }
                var typeCriteria = criteria as TypeCriteria;
                if (typeCriteria != null)
                {
                    _chartSettings.Type = typeCriteria.Type;
                }
                var scaleCriteria = criteria as ScaleCriteria;
                if (scaleCriteria != null)
                {
                    _chartSettings.Scale = scaleCriteria.Scale;
                }
                var sizeCriteria = criteria as SizeCriteria;
                if (sizeCriteria != null)
                {
                    _chartSettings.ImageSize = sizeCriteria.Size;
                }
                var movingAvgCriteria = criteria as MovingAvgCriteria;
                if (movingAvgCriteria != null)
                {
                    _chartSettings.MovingAverages.Add(movingAvgCriteria.MovingAverageInterval);
                }
                var emaCriteria = criteria as EMACriteria;
                if (emaCriteria != null)
                {
                    _chartSettings.ExponentialMovingAverages.Add(emaCriteria.EMA);
                }
                var indicatorCriteria = criteria as IndicatorCriteria;
                if (indicatorCriteria != null)
                {
                    _chartSettings.TechnicalIndicators.Add(indicatorCriteria.Indicator);
                }
                var overlaysCriteria = criteria as OverlaysCriteria;
                if (overlaysCriteria != null)
                {
                    _chartSettings.ChartOverlays.Add(overlaysCriteria.Overlay);
                }
                var compareVsCriteria = criteria as CompareVsCriteria;
                if (compareVsCriteria != null)
                {
                    _chartSettings.ComparingIDs.AddRange(compareVsCriteria.CompareVsIds);
                }
            }
        }
    }
}
