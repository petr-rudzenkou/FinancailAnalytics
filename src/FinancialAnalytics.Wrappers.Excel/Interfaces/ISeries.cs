using System;
using FinancialAnalytics.Wrappers.Office.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ISeries : IEntityWrapper<ISeries>
    {
        object NativeValues { get; set; }

        object NativeXValues { get; set; }

        bool HasDataLabels { get; set; }

        Array Values { get; set; }

        Array XValues { get; set; }

		string Name { get; set; }

		string Formula { get; set; }

        string FormulaLocal { get; }

        object Select();

		void Delete();
		
		void Copy();

		void Paste();

        IPoints Points(object index);

        IBorder Border { get; }

        IChartFillFormat Fill { get; }

        IInterior Interior { get; }

        ChartType ChartType { get; }

        int MarkerBackgroundColor { get; set; }

		int MarkerForegroundColor { get; set; }

		MarkerStyle MarkerStyle { get; set; }

        ColorIndex MarkerBackgroundColorIndex { get; set; }
		ColorIndex MarkerForegroundColorIndex { get; set; }

		int MarkerSize { get; set; }

		bool HasUndefinedType { get; }
    }
}
