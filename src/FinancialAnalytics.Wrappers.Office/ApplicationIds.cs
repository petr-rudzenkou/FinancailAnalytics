using System;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Office
{

    /// <summary>
    /// Used for getting application id
    /// </summary>
    public class ApplicationIds : Interfaces.IApplicationIds
    {
        private readonly OfficeVersion _version;

        public ApplicationIds(OfficeVersion version)
        {
            _version = version;
        }

        public OfficeVersion CurrentVersion
        {
            get { return _version; }
        }

        /// <summary>
        /// Gets Excel id
        /// </summary>
        /// <returns>String representation of id</returns>
        public string GetApplicationId()
        {
            switch (_version)
            {
                case OfficeVersion.Office2003:
                    return "11";
                case OfficeVersion.Office2007:
                    return "12";
                case OfficeVersion.Office2010:
                    return "14";
                case OfficeVersion.Office2013:
                    return "15";
                default:
                    return String.Empty;

            }            
        }
    }
}
