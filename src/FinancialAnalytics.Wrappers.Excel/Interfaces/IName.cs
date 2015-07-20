using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IName : IEntityWrapper<IName>
    {
        string RangeName { get; }

        string RangeValue { get; }

        IRange RefersToRange { get; }

        string NameLocal { get; }

        string RefersTo { get; }

        bool Visible { get; set; }

    	void Delete();
    }
}
