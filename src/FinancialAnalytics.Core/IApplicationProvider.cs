using System;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Core
{
    public interface IApplicationProvider : IDisposable
    {
        IApplication Application { get; }
        void SetApplication(IApplication application);

        void WhenReady(Action<IApplication> action);

        bool IfReady(Action<IApplication> action, bool showNotAvaliableMessage = true);

        bool IsReady();

        IDisposable SupressUpdating();

        T GetIfReady<T>(
            Func<IApplication, T> getValue,
            T defaultValue = default(T),
            bool showNotAvaliableMessage = true);

        T GetIfRangeSelected<T>(
            Func<IRange, T> getResult,
            T defaultResult = default(T));
    }
}
