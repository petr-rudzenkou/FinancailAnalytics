using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IListObject :IEntityWrapper<IListObject>
    {
        IRange Range { get; }

        IRange DataBodyRange { get; }

        IRange HeaderRowRange { get; }

        IListRows ListRows { get; }

        string Name { get; }

        bool Active { get; }
    }
}
