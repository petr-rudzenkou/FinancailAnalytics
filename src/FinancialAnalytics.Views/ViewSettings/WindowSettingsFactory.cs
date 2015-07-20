namespace FinancialAnalytics.Views.ViewSettings
{
    public class WindowSettingsFactory : IWindowSettingsFactory
    {
        public WindowSettings GetWindowSettings(ViewType viewType)
        {
            WindowSettings windowSettings;
            switch (viewType)
            {
                case ViewType.StockScreener:
                    windowSettings = new WindowSettings()
                    {
                        Width = 870,
                        Height = 700,
                    };
                    break;
                case ViewType.Portfolio:
                    windowSettings = new WindowSettings()
                    {
                        Width = 1000,
                        Height = 510,
                    };
                    break;
                case ViewType.HistoricalData:
                    windowSettings = new WindowSettings()
                    {
                        Width = 630,
                        Height = 610,
                    };
                    break;
                case ViewType.LeagueTable:
                    windowSettings = new WindowSettings()
                    {
                        Width = 500,
                        Height = 600,
                    };
                    break;
                case ViewType.Charts:
                    windowSettings = new WindowSettings()
                    {
                        Width = 860,
                        Height = 650,
                    };
                    break;
                case ViewType.Quotes:
                    windowSettings = new WindowSettings()
                    {
                        Width = 1000,
                        Height = 570,
                    };
                    break;
                case ViewType.Options:
                    windowSettings = new WindowSettings()
                    {
                        Width = 400,
                        Height = 180,
                    };
                    break;
                case ViewType.Search:
                    windowSettings = new WindowSettings()
                    {
                        Width = 300,
                        Height = 150,
                    };
                    break;
                case ViewType.Login:
                    windowSettings = new WindowSettings()
                    {
                        Width = 300,
                        Height = 250,
                    };
                    break;
                case ViewType.XChangeRates:
                    windowSettings = new WindowSettings()
                    {
                        Width = 600,
                        Height = 270,
                    };
                    break;
                default:
                    windowSettings = new WindowSettings()
                    {
                        Width = 870,
                        Height = 700,
                    };
                    break;
            }
            return windowSettings;
        }
    }
}
