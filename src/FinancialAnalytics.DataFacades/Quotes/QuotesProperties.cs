namespace FinancialAnalytics.DataFacades.Quotes
{
    public class QuotesProperties
    {
        public readonly static string[] DefaultQuoteProperties = 
        {
            QuoteProperty.Symbol.ToString(),
            QuoteProperty.Name.ToString(),
            QuoteProperty.LastTradeDate.ToString(),
            QuoteProperty.LastTradeTime.ToString(),
            QuoteProperty.Open.ToString(),
            QuoteProperty.MarketCapitalization.ToString(),
            QuoteProperty.Bid.ToString(),
            QuoteProperty.Ask.ToString(),
            QuoteProperty.PreviousClose.ToString(),
            QuoteProperty.Change.ToString(),
            QuoteProperty.PercentChange.ToString(),
            QuoteProperty.DaysLow.ToString(),
            QuoteProperty.DaysHigh.ToString(),
            QuoteProperty.Volume.ToString(),
            QuoteProperty.AverageDailyVolume.ToString(),
            QuoteProperty.EBITDA.ToString(),
            QuoteProperty.Currency.ToString(),
            QuoteProperty.StockExchange.ToString()
        };
    }
}
