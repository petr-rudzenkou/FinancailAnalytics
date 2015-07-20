using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class PriceBookRatioCriteria : RangeCriteria
    {
        public PriceBookRatioCriteria()
        {
            Name = QuoteProperty.PriceBook.ToString();
            DisplayName = "Price/Book Ratio";
        }
    }
}
