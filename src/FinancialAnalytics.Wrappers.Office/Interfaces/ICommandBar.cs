using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface ICommandBar : IEntityWrapper<ICommandBar>
	{
		ICommandBarControl FindControl(ControlType type, object id, object tag, object visible, bool recursive);

		void Reset();

		void Delete();

		bool Enabled { get; set; }

		bool Visible { get; set; }

		ICommandBarControls Controls { get; }

		int Id { get; }
	}
}