using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ICustomProperty : IEntityWrapper<ICustomProperty>
    {
        string Name { get; }

        string Value { get; set; }

        void Delete();
    }
}
