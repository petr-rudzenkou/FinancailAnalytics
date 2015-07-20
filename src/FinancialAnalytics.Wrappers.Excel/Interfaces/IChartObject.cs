using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IChartObject : IEntityWrapper<IChartObject>
    {
        IChart Chart { get; }

        string Name { get; }

        double Height { get; set; }

        double Width { get; set; }

		double Left { get; set; }

		double Top { get; set; }
		
		void Select(bool replace);

        void Activate();

        void Delete();

    }
}
