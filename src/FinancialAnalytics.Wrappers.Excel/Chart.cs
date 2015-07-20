using System;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.EventsRouting;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.ExceptionHandling;
using IChartArea = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartArea;
using IChartGroups = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartGroups;
using IChartObject = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartObject;
using IChartTitle = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartTitle;
using ILegend = FinancialAnalytics.Wrappers.Excel.Interfaces.ILegend;
using IPageSetup = FinancialAnalytics.Wrappers.Excel.Interfaces.IPageSetup;
using IPlotArea = FinancialAnalytics.Wrappers.Excel.Interfaces.IPlotArea;
using IRange = FinancialAnalytics.Wrappers.Excel.Interfaces.IRange;
using ISeriesCollection = FinancialAnalytics.Wrappers.Excel.Interfaces.ISeriesCollection;
using IShapes = FinancialAnalytics.Wrappers.Excel.Interfaces.IShapes;
using System.Reflection;

namespace FinancialAnalytics.Wrappers.Excel
{
	internal class Chart : ExcelEntityWrapper<IChart>, IChart
	{
		private Microsoft.Office.Interop.Excel.Chart _excelChart;

		public Chart(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Chart chart)
			: base(entityResolver)
		{
			if (chart == null)
				throw new ArgumentNullException("chart");
			_excelChart = chart;
			InitializeEventHandlers();
		}

		~Chart()
		{
			try
			{
				Dispose();
			}
			catch (Exception)
			{
			}
		}

        public void Calculate()
        { }

		#region Disposable pattern

		private bool disposed = false;
		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				if (disposing)
				{
					//Here we must dispose managed resources
				}
				//Here we must dispose unmanaged resources and LOH objects
				RemoveEventHandlers();
				ComObjectsFinalizer.ReleaseComObject(_excelChart);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

		

		public string Name
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelChart.Name;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelChart.Name = value;
				}
			}
		}

		public bool HasTitle
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelChart.HasTitle;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelChart.HasTitle = value;
				}
			}
		}

		public bool IsInplace
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return GetIsInplace();
				}
			}
		}

		private bool GetIsInplace()
		{
			using (new EnUsCultureInvoker())
			{
				return (Workbook == null || Workbook.IsInplace);
			}
		}

		public string CodeName
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					//if sheetcode is empty
					if (string.IsNullOrEmpty(_excelChart.CodeName))
					{
						//we need to access sheetcode via VBComponents and Excel will generate it automatically
						var workbook = _excelChart.Parent as Microsoft.Office.Interop.Excel.Workbook;
						if (workbook != null)
						{
							try
							{
								return workbook.VBProject.VBComponents.Item(Name).Name;
							}
							catch (Exception exc)
							{
								bool rethrow = ExceptionHandler.HandleException(exc);
								if (rethrow)
									throw;
							}
							finally
							{
								ComObjectsFinalizer.ReleaseComObject(workbook);
							}
						}
					}
					return _excelChart.CodeName;
				}
			}
		}

		public object SheetObject
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelChart;
				}
			}
		}

		public IChartObject ChartObject
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartObject(_excelChart.Parent as Microsoft.Office.Interop.Excel.ChartObject);
				}
			}
		}


		public IChartArea ChartArea
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartArea(_excelChart.ChartArea);
				}
			}
		}

		public IChartTitle ChartTitle
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartTitle(_excelChart.ChartTitle);
				}
			}
		}

		public IPlotArea PlotArea
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolvePlotArea(_excelChart.PlotArea);
				}
			}
		}

		public object Parent
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return GetParent();
				}
			}
		}

		public IShapes Shapes
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveShapes(_excelChart.Shapes);
				}
			}
		}

		public IApplication Application
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveApplication();
				}
			}
		}

		public ISeriesCollection SeriesCollection(object index)
		{
			using (new EnUsCultureInvoker())
			{
				return EntityResolver.ResolveSeriesCollection(
						_excelChart.SeriesCollection(index) as
						Microsoft.Office.Interop.Excel.SeriesCollection);
			}
		}

		public DisplayBlanksAs DisplayBlanksAs
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlDisplayBlanksAsToDisplayBlanksAsConverter.ConvertBack(_excelChart.DisplayBlanksAs);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelChart.DisplayBlanksAs = XlDisplayBlanksAsToDisplayBlanksAsConverter.ConvertBack(value);
				}
			}
		}

		public ChartType ChartType
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlChartTypeToChartTypeConverter.Convert(_excelChart.ChartType);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelChart.ChartType = XlChartTypeToChartTypeConverter.ConvertBack(value);
				}
			}
		}

		public void Activate()
		{
			using (new EnUsCultureInvoker())
			{
				(_excelChart as _Chart).Activate();
			}
		}

		public void Select(bool replace)
		{
			using (new EnUsCultureInvoker())
			{
				(_excelChart as _Chart).Select(replace);
			}
		}

		public void Paste()
		{
			using (new EnUsCultureInvoker())
			{
				_excelChart.Paste(Type.Missing);
			}
		}

		public void Paste(PasteToChartType pasteToChartType)
		{
			using (new EnUsCultureInvoker())
			{
				_excelChart.Paste(pasteToChartType);
			}
		}

		public bool Export(string fileFullName, string fileFormat)
		{
			using (new EnUsCultureInvoker())
			{
				return _excelChart.Export(fileFullName, fileFormat, Type.Missing);
			}
		}

		public IPageSetup PageSetup
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolvePageSetup(_excelChart.PageSetup);
				}
			}
		}

		public SheetVisibility Visible
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlSheetVisibilityToSheetVisibilityConverter.Convert(_excelChart.Visible);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelChart.Visible = XlSheetVisibilityToSheetVisibilityConverter.ConvertBack(value);
				}
			}
		}

		public void Delete()
		{
			using (new EnUsCultureInvoker())
			{
				_excelChart.Delete();
				Dispose();
			}
		}

		public void SetSourceData(IRange source)
		{
			using (new EnUsCultureInvoker())
			{
				_excelChart.SetSourceData((Microsoft.Office.Interop.Excel.Range)source.RangeObject, Type.Missing);
			}
		}

		public void SetSourceData(IRange source, RowCol plotBy)
		{
			using (new EnUsCultureInvoker())
			{
				Microsoft.Office.Interop.Excel.XlRowCol xlRowCol = XlRowColToPlotWayConverter.ConvertBack(plotBy);
				_excelChart.SetSourceData((Microsoft.Office.Interop.Excel.Range)source.RangeObject, xlRowCol);
			}
		}

		public IChart Location(ChartLocation location, string sheetName)
		{
			using (new EnUsCultureInvoker())
			{
				Microsoft.Office.Interop.Excel.XlChartLocation xlChartLocation =
					XlChartLocationToChartLocationConverter.ConvertBack(location);

				Microsoft.Office.Interop.Excel.Chart newExcelChart;
				if (string.IsNullOrEmpty(sheetName))
				{
					newExcelChart = _excelChart.Location(xlChartLocation, Type.Missing);
				}
				else
				{
					newExcelChart = _excelChart.Location(xlChartLocation, sheetName);
				}

				return EntityResolver.ResolveChart(newExcelChart);
			}
		}

		public IChart Location(ChartLocation location)
		{
			return Location(location, null);
		}

		public void CopyPicture(PictureAppearance appearance, CopyPictureFormat format, PictureAppearance size)
		{
			using (new EnUsCultureInvoker())
			{
				Microsoft.Office.Interop.Excel.XlPictureAppearance xlPictureAppearance =
					XlPictureAppearanceToPictureAppearanceConverter.ConvertBack(appearance);
				Microsoft.Office.Interop.Excel.XlCopyPictureFormat xlCopyPictureFormat =
					XlCopyPictureFormatToCopyPictureFormatConverter.ConvertBack(format);
				Microsoft.Office.Interop.Excel.XlPictureAppearance xlPictureSize =
					XlPictureAppearanceToPictureAppearanceConverter.ConvertBack(size);
				RepeatedCopyHelper.ExecuteCopyRepeated(() => _excelChart.CopyPicture(xlPictureAppearance, xlCopyPictureFormat, xlPictureSize));
			}
		}

		public object GetAxes(AxisType type)
		{
			Microsoft.Office.Interop.Excel.XlAxisType xlAxisType = XlAxisTypeToAxisTypeConverter.ConvertBack(type);
			object nativeAxes = _excelChart.Axes(xlAxisType);
			if (nativeAxes is Microsoft.Office.Interop.Excel.Axis)
			{
				return EntityResolver.ResolveAxis((Microsoft.Office.Interop.Excel.Axis)nativeAxes);
			}
			if (nativeAxes is Microsoft.Office.Interop.Excel.Axes)
			{
				return EntityResolver.ResolveAxes((Microsoft.Office.Interop.Excel.Axes)nativeAxes);
			}
			return null;
		}

		public object GetAxes(AxisType type, AxisGroup group)
		{
			Microsoft.Office.Interop.Excel.XlAxisType xlAxisType = XlAxisTypeToAxisTypeConverter.ConvertBack(type);
			Microsoft.Office.Interop.Excel.XlAxisGroup xlAxisGroup = XlAxisGroupToAxisGroupConverter.ConvertBack(group);
			object nativeAxes = _excelChart.Axes(xlAxisType, xlAxisGroup);
			if (nativeAxes is Microsoft.Office.Interop.Excel.Axis)
			{
				return EntityResolver.ResolveAxis((Microsoft.Office.Interop.Excel.Axis)nativeAxes);
			}
			if (nativeAxes is Microsoft.Office.Interop.Excel.Axes)
			{
				return EntityResolver.ResolveAxes((Microsoft.Office.Interop.Excel.Axes)nativeAxes);
			}
			return null;
		}

		protected virtual object GetParent()
		{
			using (new EnUsCultureInvoker())
			{
				object parent;
				if (_excelChart.Parent is Microsoft.Office.Interop.Excel.ChartObject)
				{
					parent =
						EntityResolver.ResolveChartObject(
							_excelChart.Parent as Microsoft.Office.Interop.Excel.ChartObject);
				}
				else
				{
					parent =
						EntityResolver.ResolveWorkbook(_excelChart.Parent as Microsoft.Office.Interop.Excel.Workbook);
				}
				return parent;
			}
		}

		public bool HasLegend
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelChart.HasLegend;
				}
			}

			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelChart.HasLegend = value;
				}
			}
		}

		public ILegend Legend
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveLegend(_excelChart.Legend);
				}
			}
		}

		public bool get_HasAxis(AxisType axisType, AxisGroup axisGroup)
		{
			using (new EnUsCultureInvoker())
			{
				return (bool)_excelChart.get_HasAxis(XlAxisTypeToAxisTypeConverter.ConvertBack(axisType), XlAxisGroupToAxisGroupConverter.ConvertBack(axisGroup));
			}

		}

		public void set_HasAxis(AxisType axisType, AxisGroup axisGroup, bool value)
		{
			using (new EnUsCultureInvoker())
			{
				_excelChart.set_HasAxis(XlAxisTypeToAxisTypeConverter.ConvertBack(axisType), XlAxisGroupToAxisGroupConverter.ConvertBack(axisGroup), (object)value);
			}

		}

		public object ChartStyle
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
                    return _excelChart.ChartStyle;
					
				}
			}

			set
			{
				using (new EnUsCultureInvoker())
				{
                    _excelChart.ChartStyle = value;
					
				}
			}
		}

		public override bool Equals(IChart obj)
		{
			using (new EnUsCultureInvoker())
			{
				if (obj == null || GetType() != obj.GetType())
				{
					return false;
				}
				Chart chart = (Chart)obj;
				try
				{
					object firstParent = _excelChart.Parent;
					object secondParent = chart._excelChart.Parent;
					bool areEquals = false;
					if (firstParent is Microsoft.Office.Interop.Excel.ChartObject &&
						secondParent is Microsoft.Office.Interop.Excel.ChartObject)
					{
						Microsoft.Office.Interop.Excel.Chart firstChart =
							(firstParent as Microsoft.Office.Interop.Excel.ChartObject).Chart;
						Microsoft.Office.Interop.Excel.Chart secondChart =
							(secondParent as Microsoft.Office.Interop.Excel.ChartObject).Chart;
						areEquals = firstChart.Equals(secondChart);
						ComObjectsFinalizer.ReleaseComObject(firstChart);
						ComObjectsFinalizer.ReleaseComObject(secondChart);
					}
					if (firstParent is Microsoft.Office.Interop.Excel.Workbook &&
						secondParent is Microsoft.Office.Interop.Excel.Workbook)
					{
						areEquals = CodeName.Equals(chart.CodeName, StringComparison.InvariantCulture) &&
									firstParent.Equals(secondParent);
					}
					ComObjectsFinalizer.ReleaseComObject(firstParent);
					ComObjectsFinalizer.ReleaseComObject(secondParent);
					return areEquals;
				}
				catch (Exception exc)
				{
					bool rethrow = ExceptionHandler.HandleException(exc);
					if (rethrow)
						throw;
				}
				return _excelChart.Equals(chart._excelChart);
			}
		}

		public bool Equals(ISheet other)
		{
			using (new EnUsCultureInvoker())
			{
				return Equals(other as IChart);
			}
		}

		public IWorkbook Workbook
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					object chartParent = null;
					Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
					Microsoft.Office.Interop.Excel.ChartObject chartObject = null;
					Microsoft.Office.Interop.Excel.Workbook workbook = null;
					try
					{
						chartParent = _excelChart.Parent;
						if (chartParent is Microsoft.Office.Interop.Excel.ChartObject)
						{
							chartObject = (Microsoft.Office.Interop.Excel.ChartObject)chartParent;
							worksheet = chartObject.Parent as Microsoft.Office.Interop.Excel.Worksheet;
							if (worksheet == null)
							{
								return null;
							}
							workbook = worksheet.Parent as Microsoft.Office.Interop.Excel.Workbook;
						}
						else if (chartParent is Microsoft.Office.Interop.Excel.Workbook)
						{
							workbook = (Microsoft.Office.Interop.Excel.Workbook)chartParent;
						}
						if (workbook == null)
						{
							return null;
						}
					}
					finally
					{
						if (chartParent != null)
						{
							ComObjectsFinalizer.ReleaseComObject(chartParent);
						}
						if (worksheet != null)
						{
							ComObjectsFinalizer.ReleaseComObject(worksheet);
						}
						if (chartObject != null)
						{
							ComObjectsFinalizer.ReleaseComObject(chartObject);
						}
					}
					return EntityResolver.ResolveWorkbook(workbook);
				}
			}
		}

		public bool HasUndefinedType
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return (int)_excelChart.ChartType == -4111;
				}
			}
		}

		public IChartGroups ChartGroups
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartGroups((Microsoft.Office.Interop.Excel.ChartGroups) _excelChart.ChartGroups());
				}
			}
		}

		public RowCol PlotBy
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlRowColToPlotWayConverter.Convert(_excelChart.PlotBy);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelChart.PlotBy = XlRowColToPlotWayConverter.ConvertBack(value);
				}
			}
		}


		public bool IsPivotChart
		{
			get 
			{
				using (new EnUsCultureInvoker())
				{ 
					return _excelChart.PivotLayout != null;
				}
			}
		}

		public void ExportAsFixedFormat(FixedFormatType formatType, string fileName)
		{
			using (new EnUsCultureInvoker())
			{
				_excelChart.GetType().InvokeMember("ExportAsFixedFormat", BindingFlags.InvokeMethod, null, _excelChart, new object[] { formatType, fileName });
			}
		}

		#region Events

		private void InitializeEventHandlers()
		{
			ChartEvents_Event chartEvents = _excelChart as ChartEvents_Event;
			if (chartEvents != null)
			{
				chartEvents.Select += OnChartSelected; 
				chartEvents.MouseDown += OnChartMouseDown;
			} 
		}

		private void RemoveEventHandlers()
		{
			ChartEvents_Event chartEvents = _excelChart as ChartEvents_Event;
			if (chartEvents != null)
			{
				chartEvents.Select -= OnChartSelected;
				chartEvents.MouseDown -= OnChartMouseDown;
			} 
		}

		private void OnChartMouseDown(int Button, int Shift, int x, int y)
		{
			if (ChartMouseDown != null)
			{
				ChartMouseDown(Button, Shift, x, y);
			}
		}

		private void OnChartSelected(int ElementID, int Arg1, int Arg2)
		{
			if (ChartSelect != null)
			{
				ChartSelect(ElementID, Arg1, Arg2); 
			}
		}

		public event ChartSelectEventHandler ChartSelect;
		public event ChartMouseDownEventHandler ChartMouseDown;

		#endregion
	}
}
