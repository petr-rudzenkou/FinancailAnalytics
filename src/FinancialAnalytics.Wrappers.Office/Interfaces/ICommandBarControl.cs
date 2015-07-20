namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface ICommandBarControl : IEntityWrapper<ICommandBarControl>
	{
		bool Enabled { get; set; }

		void Delete(object temporary);

		bool Visible { get; set; }

		string Caption { get; set; }

		string Tag { get; set; }

	    int Index { get; }

	    int Id { get; }

        int ListCount { get; }
	}
}