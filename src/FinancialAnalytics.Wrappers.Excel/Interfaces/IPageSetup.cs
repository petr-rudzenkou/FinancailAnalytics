using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPageSetup
    {
        ObjectSize ChartSize { get; set; }

        bool Equals(IPageSetup obj);

		string PrintArea { get; set; }

		object FitToPagesWide { get; set; }

		object FitToPagesTall { get; set; }

		object Zoom { get; set; }

		string LeftFooter { get; set; }

		string CenterFooter { get; set; }

		string RightFooter { get; set; }

		string LeftHeader { get; set; }

		string CenterHeader { get; set; }

		string RightHeader { get; set; }

		double TopMargin { get; set; }

		double RightMargin { get; set; }

		double BottomMargin { get; set; }

		double LeftMargin { get; set; }

		PageOrientation Orientation { get; set; }

		IApplication Application { get; }
    }
}
