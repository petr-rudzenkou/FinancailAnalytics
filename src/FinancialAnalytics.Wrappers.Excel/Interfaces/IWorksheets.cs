using System;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IWorksheets : ISheetsBase<IWorksheets, IWorksheet>
    {
        IWorksheet this[string codeName] { get; }

        Object Add(Object before, Object after, Object count, Object type);

        IWorksheet Add();
    }
}
