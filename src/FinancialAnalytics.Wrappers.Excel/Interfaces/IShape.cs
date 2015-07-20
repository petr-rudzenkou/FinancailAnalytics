using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IShape : IEntityWrapper<IShape>
    {
		int Id { get; }

    	IApplication Application { get; }

        float Height { get; set; }

        float Width { get; set; }

        ShapeType Type { get; }

		PlacementType Placement { get; set; }

	    void IncrementTop(int value);

        void Delete();

		void Select(bool replace);

        IFillFormat Fill { get; }

    	string Name { get; }

		ILineFormat Line { get; }

		void Cut();
    }
}
