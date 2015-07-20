using FinancialAnalytics.Wrappers.Excel.Converters.Localization;
using FinancialAnalytics.Wrappers.Excel.Resources;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum UnderlineStyle
    {
		[LocalizedDescription("UnderlineStyle_UnderlineStyleDouble", typeof(Content))]
        UnderlineStyleDouble,

		[LocalizedDescription("UnderlineStyle_UnderlineStyleDoubleAccounting", typeof(Content))]
        UnderlineStyleDoubleAccounting,

		[LocalizedDescription("UnderlineStyle_UnderlineStyleNone", typeof(Content))]
        UnderlineStyleNone,

		[LocalizedDescription("UnderlineStyle_UnderlineStyleSingle", typeof(Content))]
        UnderlineStyleSingle,

		[LocalizedDescription("UnderlineStyle_UnderlineStyleSingleAccounting", typeof(Content))]
        UnderlineStyleSingleAccounting
    }
}
