using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface IColorFormat
	{
		int RGB { get; set; }
		bool Equals(IColorFormat obj);
	}
}
