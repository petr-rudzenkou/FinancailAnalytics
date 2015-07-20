using FinancialAnalytics.Wrappers.Office.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IShapes : IEntitiesCollectionWrapper<IShapes, IShape>
    {
		IShape AddPicture(string fileName, bool linkToFile, bool saveWithDocument, float left, float top, float width = -1f, float height = -1f);
	    IShape AddShape(AutoShapeType shapeType, float left, float top, float width, float height);
    }
}
