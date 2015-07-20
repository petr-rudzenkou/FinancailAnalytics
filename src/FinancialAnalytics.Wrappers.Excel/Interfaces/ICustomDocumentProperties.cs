using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ICustomDocumentProperties : IEntityWrapper<ICustomDocumentProperties>
    {
        void AddProperty(string propertyName, string value);

        void DeleteProperty(string propertyName);

        string GetPropertyValue(string propertyName);
    }
}
