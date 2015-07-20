using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IRectangle
    {
        IShapeRange ShapeRange { get; }
        IInterior Interior { get; }
        IBorder Border { get; }
        IFont Font { get; }
		double Height { get; }
		double Width { get; }
    }
}
