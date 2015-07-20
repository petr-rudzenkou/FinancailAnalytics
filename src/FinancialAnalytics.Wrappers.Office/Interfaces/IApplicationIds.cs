using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
    public interface IApplicationIds
    {
        string GetApplicationId();

        OfficeVersion CurrentVersion { get; }
    }
}
