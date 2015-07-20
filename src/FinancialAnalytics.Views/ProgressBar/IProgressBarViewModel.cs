using System;
using Caliburn.Micro;

namespace FinancialAnalytics.Views.ProgressBar
{
    public interface IProgressBarViewModel : IScreen
    {
        string Caption { get; set; }

        string Title { get; set; }

        bool SupportCancellation { get; set; }

        EventHandler Cancelled { get; set; }

        void Close();
    }
}
