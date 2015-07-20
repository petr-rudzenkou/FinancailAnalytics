using System;
using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
    public class MsoShapeTypeToShapeTypeConverter
    {
        public static ShapeType Convert(MsoShapeType msoShapeType)
        {
            ShapeType shapeType = ShapeType.AutoShape;
            string shapeTypeName = msoShapeType.ToString();
            shapeTypeName = shapeTypeName.Remove(0, 3);
            if (Enum.IsDefined(typeof(ShapeType), shapeTypeName))
            {
                shapeType = (ShapeType)Enum.Parse(typeof(ShapeType), shapeTypeName, true);
            }
            return shapeType;
        }
    }
}
