using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office.Interfaces;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IShapeRange : IEntitiesCollectionWrapper<IShapeRange, IShape>
    {
        IFillFormat Fill { get; }
        ILineFormat Line { get; }
		AutoShapeType AutoShapeType { get; }
    }
}
