namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface ICommandBarPopup : ICommandBarControl
	{
		void Delete();

		ICommandBarControls Controls { get; }
	}
}