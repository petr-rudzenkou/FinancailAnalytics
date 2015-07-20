using System;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
    public interface IEntityWrapper<T> : IDisposable, IEquatable<T>
    {
    }
}
