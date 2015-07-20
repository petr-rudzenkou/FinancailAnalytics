using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IWorkbooks : IEntitiesCollectionWrapper<IWorkbooks, IWorkbook>
    {
        IWorkbook this[string fullName] { get; }
        IWorkbook Open(string fileName);
        IWorkbook Add();
        bool Contains(string fullName);
        IWorkbook Add(string templateName);
		void CloseAll();
		void CloseAllWithoutAlerts();
    }
}
