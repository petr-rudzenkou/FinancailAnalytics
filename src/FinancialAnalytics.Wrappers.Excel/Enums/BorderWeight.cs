using FinancialAnalytics.Wrappers.Excel.Converters.Localization;
using FinancialAnalytics.Wrappers.Excel.Resources;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum BorderWeight
    {
		[LocalizedDescription("BorderWeight_Hairline", typeof(Content))]
        Hairline,

		[LocalizedDescription("BorderWeight_Medium", typeof(Content))]
        Medium,

		[LocalizedDescription("BorderWeight_Thick", typeof(Content))]
        Thick,

		[LocalizedDescription("BorderWeight_Thin", typeof(Content))]
        Thin,

		[LocalizedDescription("BorderWeight_None", typeof(Content))]
        None
    }
}
