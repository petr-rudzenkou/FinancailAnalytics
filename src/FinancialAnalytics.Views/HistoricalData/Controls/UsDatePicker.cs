using System.Globalization;
using System.Reflection;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Markup;

namespace FinancialAnalytics.Views.HistoricalData.Controls
{
    public class UsDatePicker : DatePicker
    {
        public UsDatePicker()
        {
            Language = XmlLanguage.GetLanguage(new CultureInfo("en-US").IetfLanguageTag);
        }
    }
}
