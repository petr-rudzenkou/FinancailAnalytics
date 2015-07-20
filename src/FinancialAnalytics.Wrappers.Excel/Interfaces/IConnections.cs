using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IConnections : IEntitiesCollectionWrapper<IConnections, IWorkbookConnection>, IEnumerable
    {
        IWorkbookConnection Add(string name, string description, Object connectionString, Object commandText, object cmdtype);

        IWorkbookConnection this[string itemName] { get; }
    }
}
