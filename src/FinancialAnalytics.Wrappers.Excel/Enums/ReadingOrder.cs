using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
	public enum ReadingOrder
	{
		/// <summary>
		/// Depends on current OS culture. Default value.
		/// </summary>
		Context = -5002,
		/// <summary>
		/// Left to right
		/// </summary>
		LeftToRight = -5003,
		/// <summary>
		/// Right to left
		/// </summary>
		RightToLeft = -5004,
	}
}
