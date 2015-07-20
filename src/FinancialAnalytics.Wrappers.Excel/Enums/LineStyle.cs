using FinancialAnalytics.Wrappers.Excel.Converters.Localization;
using FinancialAnalytics.Wrappers.Excel.Resources;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum LineStyle
    {
		[LocalizedDescription("LineStyle_Continuous", typeof(Content))]
        Continuous = 1,

		[LocalizedDescription("LineStyle_Dash", typeof(Content))]
        Dash = -4115,

		[LocalizedDescription("LineStyle_DashDot", typeof(Content))]
        DashDot = 4,

		[LocalizedDescription("LineStyle_DashDotDot", typeof(Content))]
        DashDotDot = 5,

		[LocalizedDescription("LineStyle_Dot", typeof(Content))]
        Dot = -4118,

		[LocalizedDescription("LineStyle_Double", typeof(Content))]
        Double = -4119,

		[LocalizedDescription("LineStyle_LineStyleNone", typeof(Content))]
        LineStyleNone = 4142,

		[LocalizedDescription("LineStyle_SlantDashDot", typeof(Content))]
        SlantDashDot = 13
    }
}
