using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface IDocumentProperty : IEntityWrapper<IDocumentProperty>
	{
		string Name { get; }
		object Value { get; }
	}
}
