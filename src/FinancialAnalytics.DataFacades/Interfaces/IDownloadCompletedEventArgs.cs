using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.DataFacades.Base;

namespace FinancialAnalytics.DataFacades.Interfaces
{
    public interface IDownloadCompletedEventArgs
    {
        IResponse GetResponse();
        object UserArgs { get; }
        SettingsBase Settings { get; }
    }
}
