using System.Configuration;

namespace FinancialAnalytics.Wrappers.Office.ExceptionHandling
{
    public class ExceptionHandlingConfigurationSettings : ConfigurationSection
    {
        private const string SectionName = "ExceptionHandlerType";

        [ConfigurationProperty(SectionName)]
        public ExceptionHandlerType HandlerType
        {
            get { return (ExceptionHandlerType)base[SectionName]; }
        }
    }
}
