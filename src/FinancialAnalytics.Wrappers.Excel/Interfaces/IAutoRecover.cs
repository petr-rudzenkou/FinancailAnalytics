using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IAutoRecover : IEntityWrapper<IAutoRecover>
    {
        bool Enabled { get; set; }

        int Time { get; set; }

        string Path { get; set; }
    }
}
