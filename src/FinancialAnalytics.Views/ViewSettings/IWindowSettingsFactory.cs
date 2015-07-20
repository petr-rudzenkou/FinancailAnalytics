namespace FinancialAnalytics.Views.ViewSettings
{
    public interface IWindowSettingsFactory
    {
        WindowSettings GetWindowSettings(ViewType viewType);
    }
}
