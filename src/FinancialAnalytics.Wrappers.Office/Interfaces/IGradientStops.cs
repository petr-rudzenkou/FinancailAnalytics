using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface IGradientStops : IEntitiesCollectionWrapper<IGradientStops, IGradientStop>
	{
		void Insert(int rgb, float position, float transparency = 0f, int index = -1);
		void Delete(int index);
	}
}
