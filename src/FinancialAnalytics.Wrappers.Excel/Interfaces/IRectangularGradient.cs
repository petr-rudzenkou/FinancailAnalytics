namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IRectangularGradient : IGradient
	{
		double RectangleTop { get; set; }
		double RectangleBottom { get; set; }
		double RectangleLeft { get; set; }
		double RectangleRight { get; set; }
	}
}