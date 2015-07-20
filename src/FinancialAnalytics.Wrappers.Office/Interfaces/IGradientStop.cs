namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface IGradientStop : IEntityWrapper<IGradientStop>
	{
		IColorFormat Color { get; }
		float Position { get; set; }
		float Transparency { get; set; }
	}
}
