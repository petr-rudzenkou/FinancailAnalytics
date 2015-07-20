using System;
using System.Linq;
using System.Windows;
using Caliburn.Micro;
using FinancialAnalytics.Views.ProgressBar;
using FinancialAnalytics.Views.Screener.Base;
using FinancialAnalytics.Views.Screener.Events;
using FinancialAnalytics.Views.Screener.Interfaces;

namespace FinancialAnalytics.Views.Screener
{
    public class ScreenerViewModel : Conductor<IScreen>.Collection.OneActive, IScreenerViewModel, IHandle<RunScreenEvent>, IHandle<ScreenCompletedEvent>
    {
        private readonly IScreenBuilderViewModel _screenBuilderViewModel;
        private readonly IScreenResultsViewModel _screenResultsViewModel;

        private readonly IEventAggregator _eventAggregator;
        private readonly IProgressBarService _progressBarService;

        private readonly IScreenerResultsCollection _screenerResultsCollection;

        private IScreen _selectedView;

        public ScreenerViewModel(IScreenBuilderViewModel screenBuilderViewModel, IScreenResultsViewModel screenResultsViewModel, IEventAggregator eventAggregator, IProgressBarService progressBarService, IScreenerResultsCollection screenerResultsCollection)
        {
            _screenBuilderViewModel = screenBuilderViewModel;
            _screenResultsViewModel = screenResultsViewModel;

            _eventAggregator = eventAggregator;
            _eventAggregator.Subscribe(this);

            _progressBarService = progressBarService;
            _progressBarService.Cancelled += ScreeningCanceled;

            _screenerResultsCollection = screenerResultsCollection;

            DisplayName = Resources.ViewsResources.Screener_WindowTitle;
        }

        public IScreen SelectedView
        {
            get { return _selectedView; }
            set
            {
                if (_selectedView == value)
                    return;

                ChangeActiveItem(value, false);
            }
        }

        protected override void ChangeActiveItem(IScreen newItem, bool closePrevious)
        {
            UpdateLayout(newItem);
            base.ChangeActiveItem(newItem, closePrevious);
        }

        private void UpdateLayout(IScreen newItem)
        {
            _selectedView = newItem;
            NotifyOfPropertyChange(() => SelectedView);
        }

        protected override void OnViewAttached(object view, object context)
        {
            base.OnViewAttached(view, context);
            SetDataSource();
        }

        protected override void OnDeactivate(bool close)
        {
            base.OnDeactivate(close);
            if (close)
            {
                _eventAggregator.Publish(new ScreenerClosedEvent());
                _screenerResultsCollection.Clear();
            }
        }

        private void SetDataSource()
        {
            Items.Clear();
            Items.Add(_screenBuilderViewModel);
            Items.Add(_screenResultsViewModel);

            SelectedView = Items.FirstOrDefault();
        }


        public void Handle(RunScreenEvent message)
        {
            _progressBarService.Show(this);
        }

        public void Handle(ScreenCompletedEvent message)
        {
            _progressBarService.Close();
            if (message.HasResults)
            {
                SelectedView = _screenResultsViewModel;
            }
            else
            {
                MessageBox.Show("No results");
            }
        }

        private void ScreeningCanceled(object sender, EventArgs e)
        {
            _eventAggregator.Publish(new CancelScreenEvent());
        }
    }
}
