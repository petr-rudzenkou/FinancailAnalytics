using System.Configuration;

namespace FinancialAnalytics.Wrappers.Office.ExceptionHandling
{
    public class ExceptionHandlerType : ConfigurationElement
    {
        private const string TypePropertyName = "type";

        [ConfigurationProperty(TypePropertyName, DefaultValue = "", IsKey = true, IsRequired = true)]
        public string TypeName
        {
            get { return ((string)(base[TypePropertyName])); }
            set { base[TypePropertyName] = value; }
        }
    }
}
