using System.Runtime.InteropServices;

namespace FinancialAnalytics.Formulas
{
    [Guid("552755AA-FAC7-40A1-A4F0-A6C5B363931E")]
    //[InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [ComVisible(true)]
    public interface IFormulaProcessor
    {
        object FA(object symbols, [Optional]object dataItems, [Optional]object layout, [Optional]object destinationCell);
    }
}
