using System;
using System.Windows;
using Caliburn.Micro;
using FinancialAnalytics.Presentation.Core;
using FinancialAnalytics.Views.Screener.Events;


namespace FinancialAnalytics.Views.ProgressBar
{
    public class ProgressBarViewModel : Screen, IProgressBarViewModel
    {
        #region Fields
        private bool _supportCancellation;
        private string _caption = Resources.ViewsResources.ProgressDialog_DefaultCaption;
        private string _title = Resources.ViewsResources.ProgressDialog_DefaultTitle;
        #endregion

        #region Properties

        public string Caption
        {
            get
            {
                return _caption;
            }
            set
            {
                _caption = value;
                NotifyOfPropertyChange(() => Caption);
            }
        }

        public string Title
        {
            get
            {
                return _title;
            }
            set
            {
                _title = value;
                NotifyOfPropertyChange(() => Title);
            }
        }

        public bool SupportCancellation
        {
            get { return _supportCancellation; }
            set
            {
                _supportCancellation = value;
                NotifyOfPropertyChange(() => SupportCancellation);
            }
        }

        public EventHandler Cancelled { get; set; }

        #endregion

        public ProgressBarViewModel()
        {
            DisplayName = _title;
        }

        #region Methods

        public void Close()
        {
            if (IsActive)
                TryClose();
        }

        public void ExecuteCancel()
        {
            var handler = Cancelled;
            if (handler != null)
                handler(this, EventArgs.Empty);

            Close();
        }

        #endregion
    }
}
