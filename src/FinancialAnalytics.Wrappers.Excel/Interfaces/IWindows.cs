using System;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IWindows : IEntitiesCollectionWrapper<IWindows, IWindow>
    {
        IWindow this[String name] { get; }
    }
}
