using System;
using System.Globalization;
using System.Threading;

namespace FinancialAnalytics.Wrappers.Excel.Interception
{
    ///<summary>
    /// Temporary replaces culture of current thread to fix some problems accessing MS Office API from no US cultures.
    /// </summary>
    //NOTE: keep this in sync with its copy pastes
    public class EnUsCultureInvoker : IDisposable
    {
        private readonly CultureInfo _currentCulture;
        private static readonly CultureInfo _excelAccessCulture = new CultureInfo(1033);

        public EnUsCultureInvoker()
        {
            _currentCulture = Thread.CurrentThread.CurrentCulture;

            // improves perfromance
            if (_currentCulture.LCID != _excelAccessCulture.LCID)
            {
                Thread.CurrentThread.CurrentCulture = _excelAccessCulture;
            }
            else
            {
                _currentCulture = null;
            }

        }

        public void Dispose()
        {
            if (_currentCulture != null)
            {
                Thread.CurrentThread.CurrentCulture = _currentCulture;
            }
        }
    }

}
