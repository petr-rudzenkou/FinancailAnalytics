using System;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Vbe.Interfaces
{
	public interface ICodeModule : IEntityWrapper<ICodeModule>
    {
		void AddFromString(string codeString);
    }
}
