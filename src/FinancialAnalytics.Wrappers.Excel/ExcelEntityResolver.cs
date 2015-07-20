using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using IAxes = FinancialAnalytics.Wrappers.Excel.Interfaces.IAxes;
using IAxis = FinancialAnalytics.Wrappers.Excel.Interfaces.IAxis;
using IAxisTitle = FinancialAnalytics.Wrappers.Excel.Interfaces.IAxisTitle;
using IBorder = FinancialAnalytics.Wrappers.Excel.Interfaces.IBorder;
using IBorders = FinancialAnalytics.Wrappers.Excel.Interfaces.IBorders;
using IChartArea = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartArea;
using IChartColorFormat = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartColorFormat;
using IChartFillFormat = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartFillFormat;
using IChartObject = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartObject;
using IChartObjects = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartObjects;
using IChartTitle = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartTitle;
using ICharts = FinancialAnalytics.Wrappers.Excel.Interfaces.ICharts;
using ICustomProperties = FinancialAnalytics.Wrappers.Excel.Interfaces.ICustomProperties;
using ICustomProperty = FinancialAnalytics.Wrappers.Excel.Interfaces.ICustomProperty;
using IDataLabel = FinancialAnalytics.Wrappers.Excel.Interfaces.IDataLabel;
using IFont = FinancialAnalytics.Wrappers.Excel.Interfaces.IFont;
using IInterior = FinancialAnalytics.Wrappers.Excel.Interfaces.IInterior;
using IListObject = FinancialAnalytics.Wrappers.Excel.Interfaces.IListObject;
using IListObjects = FinancialAnalytics.Wrappers.Excel.Interfaces.IListObjects;
using IListRows = FinancialAnalytics.Wrappers.Excel.Interfaces.IListRows;
using IName = FinancialAnalytics.Wrappers.Excel.Interfaces.IName;
using INames = FinancialAnalytics.Wrappers.Excel.Interfaces.INames;
using IPageSetup = FinancialAnalytics.Wrappers.Excel.Interfaces.IPageSetup;
using IPivotField = FinancialAnalytics.Wrappers.Excel.Interfaces.IPivotField;
using IPivotFields = FinancialAnalytics.Wrappers.Excel.Interfaces.IPivotFields;
using IPivotItem = FinancialAnalytics.Wrappers.Excel.Interfaces.IPivotItem;
using IPivotTable = FinancialAnalytics.Wrappers.Excel.Interfaces.IPivotTable;
using IPlotArea = FinancialAnalytics.Wrappers.Excel.Interfaces.IPlotArea;
using IPoint = FinancialAnalytics.Wrappers.Excel.Interfaces.IPoint;
using IPoints = FinancialAnalytics.Wrappers.Excel.Interfaces.IPoints;
using IRange = FinancialAnalytics.Wrappers.Excel.Interfaces.IRange;
using ISeries = FinancialAnalytics.Wrappers.Excel.Interfaces.ISeries;
using ISeriesCollection = FinancialAnalytics.Wrappers.Excel.Interfaces.ISeriesCollection;
using IShape = FinancialAnalytics.Wrappers.Excel.Interfaces.IShape;
using IShapes = FinancialAnalytics.Wrappers.Excel.Interfaces.IShapes;
using IOLEObject = FinancialAnalytics.Wrappers.Excel.Interfaces.IOLEObject;
using IOLEObjects = FinancialAnalytics.Wrappers.Excel.Interfaces.IOLEObjects;
using ITickLabels = FinancialAnalytics.Wrappers.Excel.Interfaces.ITickLabels;
using IWindow = FinancialAnalytics.Wrappers.Excel.Interfaces.IWindow;
using IWindows = FinancialAnalytics.Wrappers.Excel.Interfaces.IWindows;
using IWorksheets = FinancialAnalytics.Wrappers.Excel.Interfaces.IWorksheets;
using IChartGroups = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartGroups;
using IChartGroup = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartGroup;
using IHiLoLines = FinancialAnalytics.Wrappers.Excel.Interfaces.IHiLoLines;
using IWorksheetFunction = FinancialAnalytics.Wrappers.Excel.Interfaces.IWorksheetFunction;

namespace FinancialAnalytics.Wrappers.Excel
{
    public class ExcelEntityResolver : EntityResolverBase
    {
        private IApplication _application;

        public IRange ResolveRange(Microsoft.Office.Interop.Excel.Range excelRange)
        {
            return excelRange == null ? null : new Range(this, excelRange);
        }

        public ExcelEntityResolver(IApplication application)
        {
            _application = application;
        }

        public ExcelEntityResolver(Microsoft.Office.Interop.Excel.Application nativeApplication)
        {
            _application = new Application(nativeApplication);
        }

        public IWorkbook ResolveWorkbook(Microsoft.Office.Interop.Excel._Workbook excelWorkbook)
        {
            return new Workbook(this, excelWorkbook);
        }

        public IShapes ResolveShapes(Microsoft.Office.Interop.Excel.Shapes excelShapes)
        {
            return new Shapes(this, excelShapes);
        }

        public FinancialAnalytics.Wrappers.Excel.Interfaces.IPanes ResolvePanes(Microsoft.Office.Interop.Excel.Panes excelPanes)
        {
            return new Panes(this, excelPanes);
        }

        public FinancialAnalytics.Wrappers.Excel.Interfaces.IPane ResolvePane(Microsoft.Office.Interop.Excel.Pane excelPane)
        {
            return new Pane(this, excelPane);
        }

		public IOLEObjects ResolveOLEObjects(Microsoft.Office.Interop.Excel.OLEObjects excelOLEObjs)
        {
			if (excelOLEObjs == null)
			{
				return null;
			}
			return new OLEObjects(this, excelOLEObjs);
        }

        public IChartObjects ResolveChartObjects(Microsoft.Office.Interop.Excel.ChartObjects excelChartObjects)
        {
            return new ChartObjects(this, excelChartObjects);
        }

        public ICustomProperties ResolveCustomProperties(Microsoft.Office.Interop.Excel.CustomProperties excelCustomProperties)
        {
            return new CustomProperties(this, excelCustomProperties);
        }

        public INames ResolveNames(Microsoft.Office.Interop.Excel.Names excelNames)
        {
            return new Names(this, excelNames);
        }

        public IName ResolveName(Microsoft.Office.Interop.Excel.Name excelName)
        {
            return new Name(this, excelName);
        }

        public IListRows ResolveListRows(Microsoft.Office.Interop.Excel.ListRows excelListRows)
        {
            return new ListRows(this, excelListRows);
        }

        public ICharts ResolveCharts(Microsoft.Office.Interop.Excel.Sheets excelCharts)
        {
            return new Charts(this, excelCharts);
        }

        public IWorksheets ResolveWorksheets(Microsoft.Office.Interop.Excel.Sheets excelWorksheets)
        {
            return new Worksheets(this, excelWorksheets);
        }

        public ICustomDocumentProperties ResolveCustomDocumentProperties(object excelCustomDocumentProperties)
        {
            return new CustomDocumentProperties(this, excelCustomDocumentProperties);
        }

        public IChartObject ResolveChartObject(Microsoft.Office.Interop.Excel.ChartObject excelChartObject)
        {
            return new ChartObject(this, excelChartObject);
        }

        public IChartArea ResolveChartArea(Microsoft.Office.Interop.Excel.ChartArea excelChartArea)
        {
            return new ChartArea(this, excelChartArea);
        }

        public IChart ResolveChart(Microsoft.Office.Interop.Excel.Chart excelChart)
        {
            return new Chart(this, excelChart);
        }

        public IListObjects ResolveListObjects(Microsoft.Office.Interop.Excel.ListObjects excelListObjects)
        {
            return new ListObjects(this, excelListObjects);
        }

        public IListObject ResolveListObject(Microsoft.Office.Interop.Excel.ListObject excelListObject)
        {
            return new ListObject(this, excelListObject);
        }

        public IApplication ResolveApplication()
        {
            return _application;
        }

        public IApplication ResolveApplication(Microsoft.Office.Interop.Excel.Application application)
        {
            return new Application(application);
        }
        
        public IWorkbooks ResolveWorkbooks(Microsoft.Office.Interop.Excel.Workbooks excelWorkbooks)
        {
            return new Workbooks(this, excelWorkbooks);
        }

        public ISheet ResolveSheet(object excelSheet)
        {
            ISheet sheet = null;
            if (excelSheet is Microsoft.Office.Interop.Excel.Worksheet)
            {
                sheet = ResolveWorksheet(excelSheet as Microsoft.Office.Interop.Excel.Worksheet);
            }
            else if (excelSheet is Microsoft.Office.Interop.Excel.Chart)
            {
                sheet = ResolveChart(excelSheet as Microsoft.Office.Interop.Excel.Chart);
            }
            return sheet;
        }

        public IWindow ResolveWindow(Microsoft.Office.Interop.Excel.Window excelWindow)
        {
            return new Window(this, excelWindow);
        }

        public ISeriesCollection ResolveSeriesCollection(Microsoft.Office.Interop.Excel.SeriesCollection excelSeriesCollection)
        {
            return new SeriesCollection(this, excelSeriesCollection);
        }

        public IPageSetup ResolvePageSetup(Microsoft.Office.Interop.Excel.PageSetup excelPageSetup)
        {
            return new PageSetup(this, excelPageSetup);
        }

        public IChartTitle ResolveChartTitle(Microsoft.Office.Interop.Excel.ChartTitle excelChartTitle)
        {
            return new ChartTitle(this, excelChartTitle);
        }

        public IFont ResolveFont(Microsoft.Office.Interop.Excel.Font excelFont)
        {
            return new Font(this, excelFont);
        }

        public IInterior ResolveInterior(Microsoft.Office.Interop.Excel.Interior excelInterator)
        {
            return new Interior(this, excelInterator);
        }

        public IChartFillFormat ResolveChartFillFormat(Microsoft.Office.Interop.Excel.ChartFillFormat excelChartFillFormat)
        {
            return new ChartFillFormat(this, excelChartFillFormat);
        }

        public IChartColorFormat ResolveChartColorFormat(Microsoft.Office.Interop.Excel.ChartColorFormat excelChartColorFormat)
        {
            return new ChartColorFormat(this, excelChartColorFormat);
        }

        public IWorksheet ResolveWorksheet(Microsoft.Office.Interop.Excel.Worksheet excelWorksheet)
        {
            return new Worksheet(this, excelWorksheet);
        }

        public IVisualBasicEditorWindow ResolveVisulaBasicEditorWindow(Microsoft.Vbe.Interop.Window vbeWindow)
        {
            return new VisualBasicEditorWindow(this, vbeWindow);
        }

        public IVisualBasicEditor ResolveVisulaBasicEditor(Microsoft.Vbe.Interop.VBE interopVisualBasicEditor)
        {
            return new VisualBasicEditor(this, interopVisualBasicEditor);
        }

        public IWindows ResolveWindows(Microsoft.Office.Interop.Excel.Windows excelWindows)
        {
            return new Windows(this, excelWindows);
        }

        public ICustomProperty ResolveCustomProperty(Microsoft.Office.Interop.Excel.CustomProperty excelCustomProperty)
        {
            return new CustomProperty(this, excelCustomProperty);
        }
        
        public ISheets ResolveSheets(Microsoft.Office.Interop.Excel.Sheets excelSheets)
        {
            return new Sheets(this, excelSheets);
        }

        public ISeries ResolveSeries(Microsoft.Office.Interop.Excel.Series excelSeries)
        {
            return new Series(this, excelSeries);
        }

		public IOLEObject ResolveOLEObject(Microsoft.Office.Interop.Excel.OLEObject excelOLEObject)
        {
			if (excelOLEObject == null)
			{
				return null;
			}
			return new OLEObject(this, excelOLEObject);
        }

        public IShape ResolveShape(Microsoft.Office.Interop.Excel.Shape excelShape)
        {
            return new Shape(this, excelShape);
        }

        public IBorder ResolveBorder(Microsoft.Office.Interop.Excel.Border excelBorder)
        {
            return new Border(this, excelBorder);
        }

        public IBorders ResolveBorders(Microsoft.Office.Interop.Excel.Borders excelBorders)
        {
            return new Borders(this, excelBorders);
        }

        public IAxis ResolveAxis(Microsoft.Office.Interop.Excel.Axis excelAxis)
        {
            return new Axis(this, excelAxis);
        }

        public IAxes ResolveAxes(Microsoft.Office.Interop.Excel.Axes excelAxes)
        {
            return new Axes(this, excelAxes);
        }

        public IAxisTitle ResolveAxisTitle(Microsoft.Office.Interop.Excel.AxisTitle excelAxisTitle)
        {
            return new AxisTitle(this, excelAxisTitle);
        }

        public IPlotArea ResolvePlotArea(Microsoft.Office.Interop.Excel.PlotArea excelPlotArea)
        {
            return new PlotArea(this, excelPlotArea);
        }

        public IDataLabel ResolveDataLabel(Microsoft.Office.Interop.Excel.DataLabel excelDataLabel)
        {
            return new DataLabel(this, excelDataLabel);
        }

        public IPoint ResolvePoint(Microsoft.Office.Interop.Excel.Point excelPoint)
        {
            return new Point(this, excelPoint);
        }

        public ITickLabels ResolveTickLabels(Microsoft.Office.Interop.Excel.TickLabels tickLabels)
        {
            return new TickLabels(this, tickLabels);
        }

        public IPoints ResolvePoints(Microsoft.Office.Interop.Excel.Points excelPoints)
        {
            return new Points(this, excelPoints);
        }

        public IPivotField ResolvePivotField(Microsoft.Office.Interop.Excel.PivotField pivotField)
        {
            return new PivotField(this, pivotField);
        }

        public IPivotItem ResolvePivotItem(Microsoft.Office.Interop.Excel.PivotItem pivotItem)
        {
            return new PivotItem(this, pivotItem);
        }

        public IPivotTable ResolvePivotTable(Microsoft.Office.Interop.Excel.PivotTable pivotTable)
        {
            return new PivotTable(this, pivotTable);
        }

        public Interfaces.IPivotCache ResolvePivotCache(Microsoft.Office.Interop.Excel.PivotCache excelPivotCache)
        {
            return new PivotCache(this, excelPivotCache);
        }

        public Interfaces.IPivotCell ResolvePivotCell(Microsoft.Office.Interop.Excel.PivotCell pivotCell)
        {
            return new PivotCell(this, pivotCell);
        }

        public object ResolvePivotTables(Microsoft.Office.Interop.Excel.PivotTables pivotTables)
        {
            return new PivotTables(this, pivotTables);
        }

        public Interfaces.IPivotCaches ResolvePivotCaches(Microsoft.Office.Interop.Excel.PivotCaches pivotCaches)
        {
            return new PivotCaches(this, pivotCaches);
        }

        public Interfaces.IWorkbookConnection ResolveWorkbookConnection(object excelWorkbookConnection)
        {
            return new WorkbookConnection(this, excelWorkbookConnection);
        }

        public ICubeField ResolveCubeField(Microsoft.Office.Interop.Excel.CubeField cubeField)
        {
            return new CubeField(this, cubeField);
        }

        public IPivotFields ResolvePivotFields(Microsoft.Office.Interop.Excel.PivotFields pivotFields)
        {
            return new PivotFields(this, pivotFields);
        }

        public Interfaces.IPivotItemList ResolvePivotItemList(Microsoft.Office.Interop.Excel.PivotItemList pivotItemList)
        {
            return new PivotItemList(this, pivotItemList);
        }

        public Interfaces.ICubeFields ResolveCubeFields(Microsoft.Office.Interop.Excel.CubeFields cubeFields)
        {
            return new CubeFields(this, cubeFields);
        }

        public Interfaces.IConnections ResolveConnections(object excelConnections)
        {
            return new Connections(this, excelConnections);
        }

        internal Interfaces.IPivotItems ResolvePivotItems(Microsoft.Office.Interop.Excel.PivotItems pivotItems)
        {
            return new PivotItems(this, pivotItems);
        }

        public Interfaces.ILegend ResolveLegend(Microsoft.Office.Interop.Excel.Legend excelLegend)
        {
            return new Legend(this, excelLegend);
        }

        public Interfaces.IAutoRecover ResolveAutoRecover(Microsoft.Office.Interop.Excel.AutoRecover excelAutoRecover)
        {
            return new AutoRecover(this, excelAutoRecover);
        }

        public Interfaces.ICOMObject ResolveCOMObject(object Target)
        {
            return new COMObjectWrapper(this, Target);
        }

        public Interfaces.IFillFormat ResolveFillFormat(Microsoft.Office.Interop.Excel.FillFormat excelFillFormat)
        {
            return new FillFormat(this, excelFillFormat);
        }

        public Interfaces.IShapeRange ResolveShapeRange(Microsoft.Office.Interop.Excel.ShapeRange excelShapeRange)
        {
            return new ShapeRange(this, excelShapeRange);
        }

        public Interfaces.IRectangle ResolveRectangle(Microsoft.Office.Interop.Excel.Rectangle excelRectangle)
        {
            return new Rectangle(this, excelRectangle);
        }

        public Interfaces.IOval ResolveOval(Microsoft.Office.Interop.Excel.Oval excelOval)
        {
            return new Oval(this, excelOval);
        }

        public Interfaces.ILineFormat ResolveLineFormat(Microsoft.Office.Interop.Excel.LineFormat excelLineFormat)
        {
            return new LineFormat(this, excelLineFormat);
        }

        public Interfaces.ILine ResolveLine(Microsoft.Office.Interop.Excel.Line excelLine)
        {
            return new Line(this, excelLine);
        }

        public Interfaces.IColorFormat ResolveColorFormat(Microsoft.Office.Interop.Excel.ColorFormat excelColorFormat)
        {
            return new ColorFormat(this, excelColorFormat);
        }

        public Interfaces.IFloor ResolveFloor(Microsoft.Office.Interop.Excel.Floor excelFloor)
        {
            return new Floor(this, excelFloor);
        }

        public Interfaces.IWalls ResolveWalls(Microsoft.Office.Interop.Excel.Walls excelWalls)
        {
            return new Walls(this, excelWalls);
        }

        public Interfaces.IDrawingObjects ResolveDrawingObjects(Microsoft.Office.Interop.Excel.DrawingObjects excelDrawingObjects)
        {
            return new DrawingObjects(this, excelDrawingObjects);
        }

        public Interfaces.IGridlines ResolveGridlines(Microsoft.Office.Interop.Excel.Gridlines excelGridlines)
        {
            return new Gridlines(this, excelGridlines);
        }

        public Interfaces.IThemeColorScheme ResolveThemeColorScheme(Microsoft.Office.Core.ThemeColorScheme excelThemeColorScheme)
        {
            return new ThemeColorScheme(this, excelThemeColorScheme);
        }

		public Interfaces.IThemeFontScheme ResolveThemeFontScheme(Microsoft.Office.Core.ThemeFontScheme excelThemeFontScheme)
        {
            return new ThemeFontScheme(this, excelThemeFontScheme);
        }

        public Interfaces.ITheme ResolveTheme(Microsoft.Office.Core.OfficeTheme excelTheme)
        {
            return new Theme(this, excelTheme);
        }

		public Interfaces.ITextBox ResolveTextBox(Microsoft.Office.Interop.Excel.TextBox excelTextBox)
		{
			return new TextBox(this, excelTextBox);
		}

		public Interfaces.ILegendEntries ResolveLegendEntries(Microsoft.Office.Interop.Excel.LegendEntries excelLegend)
		{
			return new LegendEntries(this, excelLegend);
		}

		public Interfaces.ILegendEntry ResolveLegendEntry(Microsoft.Office.Interop.Excel.LegendEntry legendEntry)
		{
			return new LegendEntry(this, legendEntry);
		}

		public Interfaces.IDataTable ResolveDataTable(Microsoft.Office.Interop.Excel.DataTable dataTable)
		{
			return new DataTable(this, dataTable);
		}

		public Interfaces.IDataLabels ResolveDataLabels(Microsoft.Office.Interop.Excel.DataLabels dataLabels)
		{
			return new DataLabels(this, dataLabels);
		}

		public Interfaces.IGroupObject ResolveGroupObject(Microsoft.Office.Interop.Excel.GroupObject groupObject)
		{
			return new GroupObject(this, groupObject);
		}

		internal Interfaces.ILegendKey ResolveLegendKey(Microsoft.Office.Interop.Excel.LegendKey legendKey)
		{
			return new LegendKey(this, legendKey);
		}

		internal Interfaces.ILinearGradient ResolveLinearGradient(object excelGradient)
		{
			return new LinearGradient(this, excelGradient);
		}

		internal Interfaces.IRectangularGradient ResolveRectangularGradient(object excelGradient)
		{
			return new RectangularGradient(this, excelGradient);
		}

		internal Interfaces.IColorStop ResolveColorStop(object excelColorStop)
		{
			return new ColorStop(this, excelColorStop);
		}

		internal Interfaces.IColorStops ResolveColorStops(object excelColorStops)
		{
			return new ColorStops(this, excelColorStops);
		}

		internal IChartGroups ResolveChartGroups(Microsoft.Office.Interop.Excel.ChartGroups excelChartGroups)
		{
			return new ChartGroups(this, excelChartGroups);
		}

		internal IChartGroup ResolveChartGroup(Microsoft.Office.Interop.Excel.ChartGroup excelChartGroup)
		{
			return new ChartGroup(this, excelChartGroup);
		}

		internal IBars ResolveUpBars(Microsoft.Office.Interop.Excel.UpBars excelBars)
		{
			return new UpBars(this, excelBars);
		}

		internal IBars ResolveDownBars(Microsoft.Office.Interop.Excel.DownBars excelBars)
		{
			return new DownBars(this, excelBars);
		}

		internal IHiLoLines ResolveHiLoLines(Microsoft.Office.Interop.Excel.HiLoLines excelHiLoLines)
		{
			return new HiLoLines(this, excelHiLoLines);
		}

		public IWorksheetFunction ResolveWorksheetFunction(Microsoft.Office.Interop.Excel.WorksheetFunction worksheetFunction)
		{
			return new WorksheetFunction(this, worksheetFunction);
		}

        public Interfaces.IValidation ResolveValidation(Microsoft.Office.Interop.Excel.Validation validation)
        {
            return new Validation(this, validation);
        }

		public Interfaces.ITextFrame2 ResolveTextFrame2(Microsoft.Office.Interop.Excel.TextFrame2 textFrame2)
        {
			return new TextFrame2(this, textFrame2);
        }

		public Interfaces.IChartFormat ResolveChartFormat(Microsoft.Office.Interop.Excel.ChartFormat chartFormat)
        {
			return new ChartFormat(this, chartFormat);
        }

        public Interfaces.ITrendline ResolveTrendline(Microsoft.Office.Interop.Excel.Trendline trendline)
        {
            return new Trendline(this, trendline);
        }

        public Interfaces.IDropLines ResolveDropLines(Microsoft.Office.Interop.Excel.DropLines dropLines)
        {
            return new DropLines(this, dropLines);
        }

        public Interfaces.IHiLoLines ResloveHiLoLines(Microsoft.Office.Interop.Excel.HiLoLines hiLoLines)
        {
            return new HiLoLines(this, hiLoLines);
        }

        public Interfaces.ISeriesLines ResolveSeriesLines(Microsoft.Office.Interop.Excel.SeriesLines seriesLines)
        {
            return new SeriesLines(this, seriesLines);
        }

        public Interfaces.IErrorBars ResolveErrorBars(Microsoft.Office.Interop.Excel.ErrorBars errorBars)
        {
            return new ErrorBars(this, errorBars);
        }

	}
}
