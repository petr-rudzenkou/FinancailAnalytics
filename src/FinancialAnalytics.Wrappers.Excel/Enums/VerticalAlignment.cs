using FinancialAnalytics.Wrappers.Excel.Converters.Localization;
using FinancialAnalytics.Wrappers.Excel.Resources;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum VerticalAlignment
    {
		[LocalizedDescription("VerticalAlignment_Top", typeof(Content))]
        Top,

		[LocalizedDescription("VerticalAlignment_Justify", typeof(Content))]
        Justify,

		[LocalizedDescription("VerticalAlignment_Distributed", typeof(Content))]
        Distributed,

		[LocalizedDescription("VerticalAlignment_Center", typeof(Content))]
        Center,

		[LocalizedDescription("VerticalAlignment_Bottom", typeof(Content))]
        Bottom
    }
}
