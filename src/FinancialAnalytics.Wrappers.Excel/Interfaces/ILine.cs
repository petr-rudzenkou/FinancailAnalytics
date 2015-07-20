using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ILine
    {
        IShapeRange ShapeRange { get; }
        IBorder Border { get; }
    }
}
