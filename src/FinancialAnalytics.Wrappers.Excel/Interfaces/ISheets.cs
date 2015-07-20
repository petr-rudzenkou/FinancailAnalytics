namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ISheets : ISheetsBase<ISheets, ISheet>
    {
		ISheet this[string codeName] { get; }
    }
}
