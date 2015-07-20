using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    /// <summary>
    /// A collection of <see cref="ICustomProperty"/> objects that represent additional information. The information can be used as metadata for XML.
    /// </summary>
    /// <seealso href="http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.customproperties.aspx"/>
    public interface ICustomProperties : IEntitiesCollectionWrapper<ICustomProperties, ICustomProperty>
    {
        ICustomProperty Add(string name, string value);
    }
}
