using System.Collections.Generic;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Core.Export
{
    public interface IDataExporter<T>
    {
        void DownInsert(IRange insertCell, string[] objHeaders, IList<T> data);
        void AcrossInsert(IRange insertCell, string[] objHeaders, IList<T> data);
    }
}
