using System.Collections.Generic;
using FinancialAnalytics.Wrappers.Office.EventsRouting;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
    public interface ICommandBars : IEntityWrapper<ICommandBars>, IEnumerable<ICommandBar>
    {
        ICommandBar this[string name] { get; }

		ICommandBarControl FindControl(object id);

        void ExecuteMso(string idMso);

		event CommandBarsOnUpdate OnUpdate;
    }
}
