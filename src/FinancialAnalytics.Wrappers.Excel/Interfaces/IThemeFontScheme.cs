using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IThemeFontScheme
	{
		void Save(string fileName);
		void Load(string fileName);
	}
}
