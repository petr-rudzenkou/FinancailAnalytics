
namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface ILineFormat
	{
		bool Visible { get; set; }
		IColorFormat ForeColor { get; }
		IColorFormat BackColor { get; }
		float Weight { get; set; }
		float Transparency { get; set; }
	}
}
