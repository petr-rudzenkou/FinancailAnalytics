using System;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface INames : IEntitiesCollectionWrapper<INames, IName>
    {
        IName Add(string name, string refersTo, bool visible, string rangeAddress);

        IName Add(
            Object name,
            Object refersTo,
            Object visible,
            Object macroType,
            Object shortcutKey,
            Object category,
            Object nameLocal,
            Object refersToLocal,
            Object categoryLocal,
            Object refersToR1C1,
            Object refersToR1C1Local);
			
		IName GetItem(object index, object indexLocal, object refersTo);
    }
}
