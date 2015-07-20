using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.EventsRouting;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IChart : IEntityWrapper<IChart>, ISheet
	{
		bool HasTitle { get; set; }

		IChartArea ChartArea { get; }

		IChartTitle ChartTitle { get; }

		IApplication Application { get; }

		ISeriesCollection SeriesCollection(object index);

		IPageSetup PageSetup { get; }

		ChartType ChartType { get; set; }

		IChartObject ChartObject { get; }

		IShapes Shapes { get; }

		object Parent { get; }
		
		void Paste();

		void Paste(PasteToChartType pasteToChartType);

		void CopyPicture(PictureAppearance appearance, CopyPictureFormat format, PictureAppearance size);

		bool Export(string fileFullName, string fileFormat);

		IChart Location(ChartLocation location);

		IChart Location(ChartLocation chartLocation, string sheetName);

		bool IsInplace { get; }

		void SetSourceData(IRange source);

		void SetSourceData(IRange source, RowCol plotBy);

		object GetAxes(AxisType type);

		object GetAxes(AxisType type, AxisGroup group);

		bool HasLegend { get; set; }

		ILegend Legend { get; }

		IPlotArea PlotArea { get; }

		DisplayBlanksAs DisplayBlanksAs { get; set; }

		bool get_HasAxis(AxisType axisType, AxisGroup axisGroup);

		void set_HasAxis(AxisType axisType, AxisGroup axisGroup, bool value);

		object ChartStyle { get; set; }

		bool HasUndefinedType { get; }

		IChartGroups ChartGroups { get; }

		RowCol PlotBy { get; set; }

		bool IsPivotChart { get; }

		void ExportAsFixedFormat(FixedFormatType formatType, string fileName);

		event ChartSelectEventHandler ChartSelect;
		event ChartMouseDownEventHandler ChartMouseDown;
	}
}
