using System;
using System.Threading.Tasks;
using System.Windows;
using Caliburn.Micro;
using FinancialAnalytics.Presentation.Core;

namespace FinancialAnalytics.Views.ProgressBar
{
    internal class ProgressBarService : IProgressBarService
    {
        private readonly IOfficeWindowManager _officeWindowManager;
        private readonly IProgressBarViewModel _progressBarViewModel;
        private string _caption = Resources.ViewsResources.ProgressDialog_DefaultCaption;
        private string _title = Resources.ViewsResources.ProgressDialog_DefaultTitle;

        public ProgressBarService(IOfficeWindowManager officeWindowManager, IProgressBarViewModel progressBarViewModel)
        {
            _officeWindowManager = officeWindowManager;
            _progressBarViewModel = progressBarViewModel;
        }

        #region IProgressBarService

        public bool IsShown { get; private set; }

        public EventHandler Cancelled { get; set; }


        public void Close()
        {
            if (_progressBarViewModel == null || !IsShown)
            {
                return;
            }
            _progressBarViewModel.Close();
            OnClose();
        }

        #endregion

        private void UserCancelled(object sender, EventArgs e)
        {
            if (Cancelled != null)
            {
                Cancelled(this, e);
            }
            OnClose();
        }

        private void OnClose()
        {
            _progressBarViewModel.Cancelled -= UserCancelled;
            IsShown = false;
        }

        public void Show(IScreen parent, string caption = null, string title = null)
        {
            _progressBarViewModel.Caption = caption ?? _caption;
            _progressBarViewModel.Title = title ?? _title;
            _progressBarViewModel.SupportCancellation = true;
            ((ProgressBarViewModel)_progressBarViewModel).Parent = parent;
            _progressBarViewModel.Cancelled += UserCancelled;

            _officeWindowManager.ShowWindow(_progressBarViewModel, window =>
            {
                window.Height = 150;
                window.Width = 370;
            });

            IsShown = true;
        }
    }
}
