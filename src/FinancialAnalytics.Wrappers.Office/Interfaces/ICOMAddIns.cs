using System.Runtime.InteropServices;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
    [ComVisible(true)]
    public interface ICOMAddIns
    {
        ICOMAddIn Item(ref object objectIndex);
    }
}
