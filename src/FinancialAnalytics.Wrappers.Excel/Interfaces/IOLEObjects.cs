using System; 
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IOLEObjects : IEntitiesCollectionWrapper<IOLEObjects, IOLEObject>
    {
        IOLEObject Add();
    }
}
