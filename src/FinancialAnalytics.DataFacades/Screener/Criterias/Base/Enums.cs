namespace FinancialAnalytics.DataFacades.Screener.Criterias.Base
{
    public enum StockScreenerCriteriaGroup
    {
        Category,
        ShareData,
        SalesAndProfitability,
        ValuationRatios,
        ResultsDisplaySettings
    }

    public enum StockScreenerProperty
    {
        Industry,
        IndexMembership,
        SharePrice,
        MarketCap,
        DividendYield,
        Beta,
        SalesRevenue,
        ProfitMargin,
        PriceEarningsRatio,
        PriceSalesRatio,
        PEGRatio,
        DisplayInfo,
    }

    public enum StockExchange
    {
        AMEX,
        NASDAQ,
        NYSE
    }
}
