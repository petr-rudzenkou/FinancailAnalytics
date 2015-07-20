namespace FinancialAnalytics.Core.Export
{
    public class DataExporterFactory : IDataExporterFactory
    {
        public IDataExporter<T> Create<T>()
        {
            return new DataExporter<T>();
        }
    }
}
