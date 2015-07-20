namespace FinancialAnalytics.Core.Export
{
    public interface IDataExporterFactory
    {
        IDataExporter<T> Create<T>();
    }
}
