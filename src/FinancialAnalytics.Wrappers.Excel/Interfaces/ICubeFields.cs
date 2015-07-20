using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ICubeFields : IEntitiesCollectionWrapper<ICubeFields, ICubeField>
    {
        ICubeField this[Object index] { get; }
    }
}
