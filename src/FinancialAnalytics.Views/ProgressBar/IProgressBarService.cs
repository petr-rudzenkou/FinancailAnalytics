using System;
using Caliburn.Micro;

namespace FinancialAnalytics.Views.ProgressBar
{
    public interface IProgressBarService
    {

        EventHandler Cancelled { get; set; }

        bool IsShown { get; }

        void Show(IScreen parent, string caption = null, string title = null);

        void Close();
    }
}
