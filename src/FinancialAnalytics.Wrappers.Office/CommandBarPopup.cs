using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	public class CommandBarPopup : CommandBarControl, ICommandBarPopup
	{
		private readonly Microsoft.Office.Core.CommandBarPopup _officeCommandBarPopup;

		public CommandBarPopup(EntityResolverBase entityResolver, Microsoft.Office.Core.CommandBarPopup commandBarPopup)
			: base(entityResolver, commandBarPopup)
		{
			_officeCommandBarPopup = commandBarPopup;
		}

		public void Reset()
		{
			_officeCommandBarPopup.Reset();
		}

		public void Delete()
		{
			_officeCommandBarPopup.Delete();
		}

		public ICommandBarControls Controls
		{
			get { return EntityResolver.ResolveCommandBarControls(_officeCommandBarPopup.Controls); }
		}
	}
}