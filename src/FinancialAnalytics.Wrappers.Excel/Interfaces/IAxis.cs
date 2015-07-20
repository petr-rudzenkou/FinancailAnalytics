using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IAxis : IEntityWrapper<IAxis>
    {
        TickLabelPosition TickLabelPosition { get; set; }

        IAxisTitle AxisTitle { get; }

        ITickLabels TickLabels { get; }

        bool HasMajorGridlines { get; set; }

        bool HasMinorGridlines { get; set; }

        bool ReversePlotOrder { get; set; }

        bool HasTitle { get; set; }

        ScaleType ScaleType { get; set; }

        CategoryType CategoryType { get; set; }

        AxisCrosses Crosses { get; set; }

        double CrossesAt { get; set; }

        double MinimumScale { get; set; }

        double MaximumScale { get; set; }

        bool MaximumScaleIsAuto { get; set; }

        bool MinimumScaleIsAuto { get; set; }

        IBorder Border { get; }

		DisplayUnit DisplayUnit { get; set; }

    	IGridlines MajorGridlines { get; }

    	IGridlines MinorGridlines { get; }

		double DisplayUnitCustom { get; set; }

		IChartFormat Format { get; }
    }
}
