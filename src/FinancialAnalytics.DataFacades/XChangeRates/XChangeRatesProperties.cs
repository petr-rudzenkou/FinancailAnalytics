namespace FinancialAnalytics.DataFacades.XChangeRates
{
    public class XChangeRatesProperties
    {
        public readonly static string[] DefaultXChangeRatesProperties = 
        {
           XChangeRatesProperty.Id.ToString(),
           XChangeRatesProperty.Name.ToString(),
           XChangeRatesProperty.Rate.ToString(),
           XChangeRatesProperty.Date.ToString(),
           XChangeRatesProperty.Time.ToString(),
           XChangeRatesProperty.Ask.ToString(),
           XChangeRatesProperty.Bid.ToString()
        };
    }
}
