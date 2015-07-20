namespace FinancialAnalytics.Views
{
    public interface IViewModelFactory
    {
        IViewModel Create(ViewType viewType);
    }
}
