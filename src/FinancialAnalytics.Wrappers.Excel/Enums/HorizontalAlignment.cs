using FinancialAnalytics.Wrappers.Excel.Converters.Localization;
using FinancialAnalytics.Wrappers.Excel.Resources;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum HorizontalAlignment
    {
		[LocalizedDescription("HorizontalAlignment_Right", typeof(Content))]
        Right,

		[LocalizedDescription("HorizontalAlignment_Left", typeof(Content))]
        Left,

		[LocalizedDescription("HorizontalAlignment_Justify", typeof(Content))]
        Justify,

		[LocalizedDescription("HorizontalAlignment_Distributed", typeof(Content))]
        Distributed,

		[LocalizedDescription("HorizontalAlignment_Center", typeof(Content))]
        Center,

		[LocalizedDescription("HorizontalAlignment_General", typeof(Content))]
        General,

		[LocalizedDescription("HorizontalAlignment_Fill", typeof(Content))]
        Fill,

		[LocalizedDescription("HorizontalAlignment_CenterAcrossSelection", typeof(Content))]
        CenterAcrossSelection
    }
}
