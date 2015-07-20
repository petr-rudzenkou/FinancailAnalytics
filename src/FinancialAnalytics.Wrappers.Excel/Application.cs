using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.EventsRouting;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Converters;
using FinancialAnalytics.Wrappers.Office.Enums;
using FinancialAnalytics.Wrappers.Office.ExceptionHandling;
using FinancialAnalytics.Wrappers.Office.Interfaces;
using FinancialAnalytics.Wrappers.Office.Windows;
using MSExcel = Microsoft.Office.Interop.Excel;
using MSOffice = Microsoft.Office.Core;

namespace FinancialAnalytics.Wrappers.Excel
{
    //TODO: refractor - split into 2 classes - Application and ExternalApplication
    // * NDepend shows that CO has only one critical code issue - it is extremely big class Application because of ExternalApplication logic inside
    // * 3 bugs were introduces and resolved by 4 developers taking not less then 10 man days (1 of Linking, 2 of CADA) because of ExternalApplication logic inside
    // * initialization of hosted in Excel Application degrades start up performance (e.g. Process.GetCurrentProcess)  because of ExternalApplication logic inside
    // * other issues also could exist (e.g. hardening debug, degraded performance, bad hacks for COM threading related stuff)  because of ExternalApplication logic inside
    public class Application : IApplication
    {
        #region Fields

        internal const string EXCEL_APPLICATION_EVENTS_INTERFACE_GUID = "00024413-0000-0000-C000-000000000046";
    
        private static readonly object _locker = new object();
        private readonly ExcelEntityResolver _entityResolver;
        private readonly bool _isStartedWithoutInitialization;
        private readonly ExcelProcessManager _excelProcessManager;
        private Microsoft.Office.Interop.Excel.Application _excelApplication;
        private bool _autoLaunch = false;
        private bool _hidden = false;
        private IApplicationIds _applicationIds;
        private ApplicationEvents _applicationEvents;

        #endregion

        #region Constructors

        public Application(Microsoft.Office.Interop.Excel.Application application)
            : this()
        {
            if (application == null)
                throw new ArgumentNullException("application");
            _excelApplication = application;
            _applicationIds = null;
            _applicationEvents.Attach(application, _entityResolver);
        }

        public Application(IApplicationIds applicationIds)
            : this()
        {
            _applicationIds = applicationIds;
            _isStartedWithoutInitialization = true;
        }

        private Application()
        {
            _applicationIds = new ApplicationIds(OfficeVersion.Other);
            _entityResolver = new ExcelEntityResolver(this);
            //ApplicationEvents.EntityResolver = _entityResolver;
            //ApplicationEvents.ExcelApplication = this;
            _applicationEvents = new ApplicationEvents();
            _excelProcessManager = new ExcelProcessManager();
            
        }

        #endregion

        #region IDisposable implementation

        private bool disposed; // to detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                disposed = true;
                if (disposing)
                {
                    // Here we must release managed resources.
                }
                // Here we must Release unmanaged resources.
                // Set large fields to null. For clear LOH
                ReleaseExcelApplication();
            }
        }

        private void ReleaseExcelApplication()
        {
            lock (_locker)
            {
                _applicationEvents.Deattach();
                if (_excelApplication != null)
                {
                    ComObjectsFinalizer.FinalReleaseComObject(_excelApplication);
                    _excelApplication = null;
                }
            }
        }

        [HandleProcessCorruptedStateExceptions]
        private void ReleaseExcelApplicationSafely()
        {
            lock (_locker)
            {
                try
                {
                    _applicationEvents.Deattach();
                    if (_excelApplication != null)
                    {
                        ComObjectsFinalizer.FinalReleaseComObject(_excelApplication);
                    }
                }
                catch (AccessViolationException)
                {
                    //do nothing because it's already dead application
                }
                finally
                {
                    _excelApplication = null;
                }
            }
        }

        ~Application()
        {
            try
            {
                Dispose(false);
            }
            catch (Exception)
            {
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion

        #region Implementation

        public void Calculate()
        {
            using (new EnUsCultureInvoker())
            {
                _excelApplication.Calculate();
            }
        }

        private IApplicationIds ApplicationIds
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    if (_applicationIds == null)
                    {
                        if (_excelApplication != null)
                        {
                            OfficeVersion officeVersion =
                                ApplicationVersionToOfficeVersionConverter.Convert(_excelApplication);
                            _applicationIds = new ApplicationIds(officeVersion);
                        }
                        else
                        {
                            //don't store into field (only for prevent NullReferenceExceptions)
                            return new ApplicationIds(OfficeVersion.Other);
                        }
                    }
                    return _applicationIds;
                }
            }
        }

        private string ExcelPath
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return Path.Combine(OfficePathHelper.GetOfficePath(ApplicationIds.GetApplicationId()), "excel");
                }
            }
        }

        internal void StartExcelProcess()
        {
            ReleaseExcelApplication();
            if (_hidden)
            {
                StartHiddenExcelApplication();
                return;
            }
            if (ApplicationIds.CurrentVersion == OfficeVersion.Other)
            {
                StartDefaultExcelApplication();
                return;
            }

            using (new EnUsCultureInvoker())
            {
                ProcessStartInfo processStartInfo = new ProcessStartInfo(ExcelPath);
                string applicationId = ApplicationIds.GetApplicationId();
                if (applicationId.Equals("15", StringComparison.InvariantCultureIgnoreCase))
                {
                    processStartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                }

                //Excel 2013 need to be started focused, and focus removed after that - only then it is registered in ROT
                //Remember current active window
                IntPtr currentWindowPointer = IntPtr.Zero;
                if (applicationId.Equals("15", StringComparison.InvariantCultureIgnoreCase))
                {
                    currentWindowPointer = GetForegroundWindow();
                }

                Process excelProcess = Process.Start(processStartInfo);

                while (_excelApplication == null)
                {
                    System.Threading.Thread.Sleep(50);

                    if (excelProcess == null) // shouldn't occur
                        throw new InvalidOperationException("Excel process is reused");

                    if (excelProcess.HasExited) // user has quickly closed Excel
                        throw new InvalidOperationException("Excel process is terminated");

                    //Restore active window - at this time started Excel 2013 is focused, focus need to be removed, only then Excel is registered in ROT
                    if (applicationId.Equals("15", StringComparison.InvariantCultureIgnoreCase) && currentWindowPointer != IntPtr.Zero && GetForegroundWindow() != currentWindowPointer)
                    {
                        SetForegroundWindow(currentWindowPointer);
                    }

                    IDictionary<MSExcel.Application, uint> map = _excelProcessManager.GetApplications();

                    _excelApplication = map.FirstOrDefault(kvp => kvp.Value == (uint)excelProcess.Id).Key;
                }

                System.Threading.Thread.CurrentThread.Join(500);

                // fix: now Excel is not accessible as COM server,
                // to fix it we should move focus from it, or minimize it. Excel 2013 has slightly different behavior
                if (!applicationId.Equals("15", StringComparison.InvariantCultureIgnoreCase))
                {
                    _excelApplication.WindowState = MSExcel.XlWindowState.xlMinimized;
                }

                while (!IsApplicationValid())
                {
                    System.Threading.Thread.CurrentThread.Join(500);
                }

                _applicationEvents.Attach(_excelApplication, _entityResolver);
            }
        }

        private void StartDefaultExcelApplication()
        {
            using (new EnUsCultureInvoker())
            {
                _excelApplication = new Microsoft.Office.Interop.Excel.Application { Visible = true };
                _excelApplication.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized;
                _excelApplication.Workbooks.Add();
                System.Threading.Thread.Sleep(200);
                _applicationEvents.Attach(_excelApplication, _entityResolver);
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern int SetForegroundWindow(IntPtr hwnd);

        private void StartHiddenExcelApplication()
        {
            using (new EnUsCultureInvoker())
            {
                ProcessStartInfo processStartInfo = new ProcessStartInfo(ExcelPath);
                //for Office 2010 there is no OnDisconnection if we start with WindowStyle.Hidden
                //Comment from IZ: 
                //"...cannot be removed, as it leads to crashes during hidden update/redirect in Linking"
                string applicationId = ApplicationIds.GetApplicationId();
                if (applicationId.Equals("14", StringComparison.InvariantCultureIgnoreCase) || applicationId.Equals("15", StringComparison.InvariantCultureIgnoreCase))
                {
                    processStartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                }
                else
                {
                    processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                }


                //Excel 2013 need to be started focused, and focus removed after that - only then it is registered in ROT
                //Remember current active window
                IntPtr currentWindowPointer = IntPtr.Zero;
                if (applicationId.Equals("15", StringComparison.InvariantCultureIgnoreCase))
                {
                    currentWindowPointer = GetForegroundWindow();
                }

                processStartInfo.UseShellExecute = true;
                processStartInfo.Arguments = "/linkingonly";
                Process excelProcess = Process.Start(processStartInfo);

                while (_excelApplication == null)
                {
                    System.Threading.Thread.Sleep(50);

                    if (excelProcess == null) // shouldn't occur
                        throw new InvalidOperationException("Excel process is reused");

                    if (excelProcess.HasExited) // user has quickly closed Excel
                        throw new InvalidOperationException("Excel process is terminated");

                    //Restore active window - at this time started Excel 2013 is focused, focus need to be removed, only then Excel is registered in ROT
                    if (applicationId.Equals("15", StringComparison.InvariantCultureIgnoreCase) && currentWindowPointer != IntPtr.Zero && GetForegroundWindow() != currentWindowPointer)
                    {
                        SetForegroundWindow(currentWindowPointer);
                    }

                    IDictionary<MSExcel.Application, uint> map = _excelProcessManager.GetApplications(true);

                    _excelApplication = map.FirstOrDefault(kvp => kvp.Value == (uint)excelProcess.Id).Key;
                    //_excelApplication.Visible = false;
                }

                System.Threading.Thread.CurrentThread.Join(500);

                // fix: now Excel is not accessible as COM server,
                // to fix it we should move focus from it, or minimize it
                //_excelApplication.WindowState = MSExcel.XlWindowState.xlMinimized;

                while (!IsApplicationValid())
                {
                    System.Threading.Thread.CurrentThread.Join(500);
                }

                _applicationEvents.Attach(_excelApplication, _entityResolver);

            }
        }

        private bool IsApplicationStarted()
        {
            using (new EnUsCultureInvoker())
            {
                try
                {
                    if (!IsApplicationValid())
                    {
                        Microsoft.Office.Interop.Excel.Application excelApplication = _excelProcessManager.GetActive(ApplicationIds.GetApplicationId(), AllowHidden);
                        if (excelApplication == null)
                        {
                            return false;
                        }
                        ComObjectsFinalizer.ReleaseComObject(excelApplication);
                        return true;
                    }
                }
                catch (Exception)
                {
                    //Fix for #372738 and #372377 (commented exception hanlding) - Exception handler unloaded before closing of PowerPoint
                    //bool rethrow = ExceptionPolicy.HandleException(ex, PolicyNames.WrappersPolicy);
                    //if (rethrow)
                    //    throw;
                    ReleaseExcelApplication();
                    return false;
                }
                return true;
            }
        }

        protected virtual object GetSelection()
        {
            using (new EnUsCultureInvoker())
            {
                object selectionObject = null;
                try
                {
                    object nativeSelection = _excelApplication.Selection;
                    if (nativeSelection != null)
                    {
                        if (nativeSelection is Microsoft.Office.Interop.Excel.Range)
                        {
                            selectionObject =
                                _entityResolver.ResolveRange(nativeSelection as Microsoft.Office.Interop.Excel.Range);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.ChartArea)
                        {
                            selectionObject =
                                _entityResolver.ResolveChartArea(nativeSelection as Microsoft.Office.Interop.Excel.ChartArea);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.ChartObject)
                        {
                            selectionObject =
                                _entityResolver.ResolveChartObject(nativeSelection as Microsoft.Office.Interop.Excel.ChartObject);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Rectangle)
                        {
                            selectionObject =
                                _entityResolver.ResolveRectangle(nativeSelection as Microsoft.Office.Interop.Excel.Rectangle);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Oval)
                        {
                            selectionObject =
                                _entityResolver.ResolveOval(nativeSelection as Microsoft.Office.Interop.Excel.Oval);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Line)
                        {
                            selectionObject =
                                _entityResolver.ResolveLine(nativeSelection as Microsoft.Office.Interop.Excel.Line);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Shape)
                        {
                            selectionObject =
                                _entityResolver.ResolveShape(nativeSelection as Microsoft.Office.Interop.Excel.Shape);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.PlotArea)
                        {
                            selectionObject =
                                _entityResolver.ResolvePlotArea(nativeSelection as Microsoft.Office.Interop.Excel.PlotArea);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Legend)
                        {
                            selectionObject =
                                _entityResolver.ResolveLegend(nativeSelection as Microsoft.Office.Interop.Excel.Legend);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Floor)
                        {
                            selectionObject =
                                _entityResolver.ResolveFloor(nativeSelection as Microsoft.Office.Interop.Excel.Floor);                            
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Walls)
                        {
                            selectionObject =
                                _entityResolver.ResolveWalls(nativeSelection as Microsoft.Office.Interop.Excel.Walls);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.DrawingObjects)
                        {
                            selectionObject =
                                _entityResolver.ResolveDrawingObjects(nativeSelection as Microsoft.Office.Interop.Excel.DrawingObjects);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Series)
                        {
                            selectionObject =
                                _entityResolver.ResolveSeries(nativeSelection as Microsoft.Office.Interop.Excel.Series);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Point)
                        {
                            selectionObject =
                                _entityResolver.ResolvePoint(nativeSelection as Microsoft.Office.Interop.Excel.Point);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Gridlines)
                        {
                            selectionObject =
                                _entityResolver.ResolveGridlines(nativeSelection as Microsoft.Office.Interop.Excel.Gridlines);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Axis)
                        {
                            selectionObject =
                                _entityResolver.ResolveAxis(nativeSelection as Microsoft.Office.Interop.Excel.Axis);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.TextBox)
                        {
                            selectionObject =
                                _entityResolver.ResolveTextBox(nativeSelection as Microsoft.Office.Interop.Excel.TextBox);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.ChartTitle)
                        {
                            selectionObject =
                                _entityResolver.ResolveChartTitle(nativeSelection as Microsoft.Office.Interop.Excel.ChartTitle);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.AxisTitle)
                        {
                            selectionObject =
                                _entityResolver.ResolveAxisTitle(nativeSelection as Microsoft.Office.Interop.Excel.AxisTitle);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.LegendEntry)
                        {
                            selectionObject =
                                _entityResolver.ResolveLegendEntry(nativeSelection as Microsoft.Office.Interop.Excel.LegendEntry);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.DataTable)
                        {
                            selectionObject =
                                _entityResolver.ResolveDataTable(nativeSelection as Microsoft.Office.Interop.Excel.DataTable);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.DataLabel)
                        {
                            selectionObject =
                                _entityResolver.ResolveDataLabel(nativeSelection as Microsoft.Office.Interop.Excel.DataLabel);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.DataLabels)
                        {
                            selectionObject =
                                _entityResolver.ResolveDataLabels(nativeSelection as Microsoft.Office.Interop.Excel.DataLabels);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.GroupObject)
                        {
                            selectionObject =
                                _entityResolver.ResolveGroupObject(nativeSelection as Microsoft.Office.Interop.Excel.GroupObject);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.OLEObject)
                        {
                            selectionObject =
                                _entityResolver.ResolveOLEObject(nativeSelection as Microsoft.Office.Interop.Excel.OLEObject);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.Trendline)
                        {
                            selectionObject =
                                _entityResolver.ResolveTrendline(nativeSelection as Microsoft.Office.Interop.Excel.Trendline);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.DropLines)
                        {
                            selectionObject =
                                _entityResolver.ResolveDropLines(nativeSelection as Microsoft.Office.Interop.Excel.DropLines);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.HiLoLines)
                        {
                            selectionObject =
                                _entityResolver.ResloveHiLoLines(nativeSelection as Microsoft.Office.Interop.Excel.HiLoLines);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.SeriesLines)
                        {
                            selectionObject =
                                _entityResolver.ResolveSeriesLines(nativeSelection as Microsoft.Office.Interop.Excel.SeriesLines);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.UpBars)
                        {
                            selectionObject =
                                _entityResolver.ResolveUpBars(nativeSelection as Microsoft.Office.Interop.Excel.UpBars);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.DownBars)
                        {
                            selectionObject =
                                _entityResolver.ResolveDownBars (nativeSelection as Microsoft.Office.Interop.Excel.DownBars);
                        }
                        else if (nativeSelection is Microsoft.Office.Interop.Excel.ErrorBars)
                        {
                            selectionObject =
                                _entityResolver.ResolveErrorBars(nativeSelection as Microsoft.Office.Interop.Excel.ErrorBars);
                        }
                    }
                }
                catch (COMException)
                {
                    return null;
                }

                return selectionObject;
            }
        }

        public object SelectionObject
        {
            get { return _excelApplication.Selection; }
        }

        #endregion

        #region Properties

        public ApplicationVersion ApplicationVersion
        {
            get
            {
                InitializeApplication();
                using (new EnUsCultureInvoker())
                {
                    return VersionStringToApplicationVersionConverter.Convert(_excelApplication.Version);
                }
            }
        }

        public CultureInfo UICulture
        {
            get
            {
                int languageId = _excelApplication.LanguageSettings.LanguageID[Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI];
                return new CultureInfo(languageId);
            }
        }

        public object ApplicationObject
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelApplication;
                }
            }
        }

        public bool UserControl
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelApplication.UserControl;
                }
            }

            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelApplication.UserControl = value;
                }
            }
        }

        public bool AutoLaunch
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _autoLaunch;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _autoLaunch = value;
                }
            }
        }

        public bool Hidden
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _hidden;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _hidden = value;
                    if (_hidden)
                    {
                        AllowHidden = true;
                    }
                }
            }
        }

        public bool IsInitializedWithApplication
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return !_isStartedWithoutInitialization;
                }
            }
        }

        public int Hinstance
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.Hinstance;
                }
            }
        }

        public long HinstancePtr
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    if (ApplicationVersion == Office.Enums.ApplicationVersion.Office2007)
                    {
                        //HinstancePtr was introduced for Office 2010
                        return Hinstance;
                    }
                    LateBindingInvoker invoker = new LateBindingInvoker(_excelApplication);
                    object value = invoker.InvokeGetPropertyValue("HinstancePtr");
					if (Marshal.SizeOf(value) == 8)
					{
						//x64 office, convert to Int64
						return Convert.ToInt64(value);
					}
					else
					{
						//x32 office, convert to Int32
						return Convert.ToInt32(value); 
					}
                }
            }
        }

        public int Hwnd
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.Hwnd;
                }
            }
        }

        public bool IsStarted
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return IsApplicationStarted();
                }
            }
        }

        public IWindows Windows
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _entityResolver.ResolveWindows(_excelApplication.Windows);
                }
            }
        }

        public IWindow ActiveWindow
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    IWindow window = null;
                    object activeWindow = _excelApplication.ActiveWindow;
                    if (activeWindow != null)
                    {
                        window = _entityResolver.ResolveWindow(_excelApplication.ActiveWindow);
                    }
                    return window;
                }
            }
        }

        public IWorkbooks Workbooks
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _entityResolver.ResolveWorkbooks(_excelApplication.Workbooks);
                }
            }
        }

        public IWorkbook ActiveWorkbook
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    IWorkbook workbook = null;
                    lock (_locker)
                    {
                        Microsoft.Office.Interop.Excel.Workbook nativeWorkbook = _excelApplication.ActiveWorkbook;
                        if (nativeWorkbook != null)
                        {
                            workbook = _entityResolver.ResolveWorkbook(nativeWorkbook);
                        }
                    }
                    return workbook;
                }
            }
        }

        public ISheet ActiveSheet
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    ISheet sheet = null;
                    object activeSheet = _excelApplication.ActiveSheet;
                    if (activeSheet != null)
                    {
                        sheet = _entityResolver.ResolveSheet(activeSheet);
                    }
                    return sheet;
                }
            }
        }

        public IChart ActiveChart
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    IChart chart = null;
                    MSExcel.Chart nativeChart = _excelApplication.ActiveChart;
                    if (nativeChart != null)
                    {
                        chart = _entityResolver.ResolveChart(nativeChart);
                    }
                    return chart;
                }
            }
        }

        public IRange ActiveCell
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    IRange cell = null;
                    MSExcel.Range nativeRange = _excelApplication.ActiveCell;
                    if (nativeRange != null)
                    {
                        cell = _entityResolver.ResolveRange(nativeRange);
                    }
                    return cell;
                }
            }
        }

        public object Selection
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return GetSelection();
                }
            }
        }

        public bool IsMultipleSelection
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    MSExcel.Areas selectionAreas = null;
                    if (_excelApplication.Selection != null)
                    {
                        selectionAreas = (MSExcel.Areas)_excelApplication.Selection.GetType()
                            .InvokeMember("Areas", System.Reflection.BindingFlags.GetProperty, null,
                                          _excelApplication.Selection, new object[] { });
                    }
                    return selectionAreas != null && selectionAreas.Count > 1;
                }
            }
        }

        public int WindowHandle
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.Hwnd;
                }
            }
        }

        public IntPtr TopWindowHandle
        {
            get
            {
                return (IntPtr)WindowHandle;
            }
        }

        public bool DisplayAlerts
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.DisplayAlerts;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    _excelApplication.DisplayAlerts = value;
                }
            }
        }

        public string Version
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.Version;
                }
            }
        }

        public object StatusBar
        {
            set
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    _excelApplication.StatusBar = value;
                }
            }
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.StatusBar;
                }
            }
        }

        public bool ScreenUpdating
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.ScreenUpdating;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    _excelApplication.ScreenUpdating = value;
                }
            }
        }

        public ReferenceStyle ReferenceStyle
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlReferenceStyleToReferenceStyleConverter.Convert(_excelApplication.ReferenceStyle);
                }
            }
        }

        public bool EnableEvents
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.EnableEvents;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    _excelApplication.EnableEvents = value;
                }
            }
        }

        public bool Visible
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.Visible;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    _excelApplication.Visible = value;
                }
            }
        }

        public WindowState WindowState
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return XlWindowStateToWindowStateConverter.Convert(_excelApplication.WindowState);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    _excelApplication.WindowState = XlWindowStateToWindowStateConverter.ConvertBack(value);
                }
            }
        }

        public Calculation Calculation
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return XlCalculationToCalculationConverter.Convert(_excelApplication.Calculation);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    _excelApplication.Calculation = XlCalculationToCalculationConverter.ConvertBack(value);
                }
            }
        }

        public ICommandBars CommandBars
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _entityResolver.ResolveCommandBars(_excelApplication.CommandBars);
                }
            }
        }

        public IVisualBasicEditor VBE
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _entityResolver.ResolveVisulaBasicEditor(_excelApplication.VBE);
                }
            }
        }

        public ICOMAddIns COMAddIns
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    ICOMAddIns addins = _entityResolver.ResolveCOMAddIns((MSOffice.COMAddIns)_excelApplication.GetType().InvokeMember("COMAddIns", BindingFlags.GetProperty, null, _excelApplication, null));
                    return addins;
                }
            }
        }

        public bool AllowHidden { get; set; }

        public bool SelectionEventsDisabled
        {
            get { return _applicationEvents.SelectionEventsDisabled; }
            set { _applicationEvents.SelectionEventsDisabled = value; }
        }

        public bool Interactive
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.Interactive;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    _excelApplication.Interactive = value;
                }
            }
        }

        private static readonly XlMousePoinerToMousePointerConverter _xlMousePoinerToMousePointerConverter = new XlMousePoinerToMousePointerConverter();

        public MousePointer Cursor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _xlMousePoinerToMousePointerConverter.Convert(this._excelApplication.Cursor);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    this._excelApplication.Cursor = _xlMousePoinerToMousePointerConverter.ConvertBack(value);
                }
            }
        }

        public IEnumerable<string> AddInNames
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    List<string> addInNameList = new List<string>();
                    for (int i = 1; i <= _excelApplication.AddIns.Count; i++)
                    {
                        MSExcel.AddIn addIn = _excelApplication.AddIns[i];
                        try
                        {
                            addInNameList.Add(addIn.Title);
                        }
                        catch (Exception) { }
                    }
                    return addInNameList;
                }
            }
        }

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.Name;
                }
            }
        }

        public IAutoRecover AutoRecover
        {
            get { return _entityResolver.ResolveAutoRecover(_excelApplication.AutoRecover); }
        }

        public bool Ready
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    if (_excelApplication == null)
                    {
                        return false;
                    }
                    return _excelApplication.Ready;
                }
            }
        }

        public bool DisplayClipboardWindow
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    return _excelApplication.DisplayClipboardWindow;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    InitializeApplication();
                    _excelApplication.DisplayClipboardWindow = value;
                }
            }
        }

        public IEnumerable<int> TopLevelWindowsHandles
        {
            get 
            {
                return GetTopLevelWindowHandles(this);
            }
        }

	    public IWorksheetFunction WorksheetFunction
	    {
		    get
		    {
			    using (new EnUsCultureInvoker())
			    {
				    return _entityResolver.ResolveWorksheetFunction(_excelApplication.WorksheetFunction);
			    }
		    }
	    }

        #endregion

        #region Methods

        public IRange GetCaller()
        {
            using (new EnUsCultureInvoker())
            {
                return _entityResolver.ResolveRange(_excelApplication.get_Caller());
            }
        }
        public object Evaluate(object value)
        {
            using (new EnUsCultureInvoker())
            {
                object result = _excelApplication.Evaluate(value);
                if (result is Microsoft.Office.Interop.Excel.Range)
                {
                    return _entityResolver.ResolveRange(result as Microsoft.Office.Interop.Excel.Range);
                }
                return result;
            }
        }

        public void Disable()
        {
            try
            {
                _excelApplication.ScreenUpdating = false;
                _excelApplication.EnableEvents = false;
            }
            catch (Exception ex)
            {
                bool rethrow = ExceptionHandler.HandleException(ex);
                if (rethrow)
                    throw;
            }
        }

        public void Enable()
        {
            try
            {
                _excelApplication.ScreenUpdating = true;
                _excelApplication.EnableEvents = true;
            }
            catch (Exception ex)
            {
                bool rethrow = ExceptionHandler.HandleException(ex);
                if (rethrow)
                    throw;
            }
        }

        public void Quit()
        {
            using (new EnUsCultureInvoker())
            {
                if (IsStarted)
                {
                    try
                    {
                        _excelApplication.Quit();
                        Dispose();
                    }
                    catch (COMException ex)
                    {
                        bool rethrow = ExceptionHandler.HandleException(ex);
                        if (rethrow)
                            throw;
                    }
                }
            }
        }

        public void InitializeApplication()
        {
            using (new EnUsCultureInvoker())
            {
                try
                {
                    if (!IsApplicationValid())
                    {
                        ReleaseExcelApplication();
                        _excelApplication = _excelProcessManager.GetActive(ApplicationIds.GetApplicationId(), AllowHidden);
                        if (_excelApplication != null)
                        {
                            //Problem in Excel 2003 with Authors broken templates was fixed with string below, commented for now, as it is allows to close changed workbook without save dialog in specific cases.
                            //_excelApplication.DisplayAlerts = false;
                            return;
                        }

                        if (_autoLaunch)
                            StartExcelProcess();
                    }
                }
                catch (Exception ex)
                {
                    bool rethrow = ExceptionHandler.HandleException(ex);
                    if (rethrow)
                        throw;
                    //start new Excel
                    if (_autoLaunch)
                    {
                        StartExcelProcess();
                    }
                }
            }
        }

        [HandleProcessCorruptedStateExceptions]
        public bool IsApplicationValid()
        {
            using (new EnUsCultureInvoker())
            {
                try
                {
                    if (_excelApplication == null)
                    {
                        return false;
                    }
                    // try to access application's property
                    // if no exception was thrown application  object in valid
#pragma warning disable 168
                    string version = _excelApplication.Version;
#pragma warning restore 168
                    return true;
                }
                catch (InvalidComObjectException)
                {
                    //InvalidComObjectException ("COM object that has been separated from its underlying RCW cannot be used.")
                    //means that application is already disposed - we should not dispose it one more time
                    //This catch is fix of StackOverflowException with CADA R4_2012 module:
                    //ReleaseExcelApplication() ->
                    //  RemoveEventsConnection() ->
                    //    IsApplicationStarted() ->
                    //      IsApplicationValid() ->
                    //        ReleaseExcelApplication() -> ...
                    return false;
                }
                catch (COMException cexc)
                {
                    //ErrorCode == -2147352571 : "Type mismatch. (Exception from HRESULT: 0x80020005 (DISP_E_TYPEMISMATCH))"
                    //throws only in Excel 2010 in case hovered [Paste] buttons in the range popup menu.
                    //But application is valid.
                    //ExceptionHandler.LogException(exc);
                    //Add ErrorCode == -2147417846 - Error The message filter indicated that the application is busy - This error is detaching all the events even excel application is valid.
                    if (cexc.ErrorCode == -2147352571 || cexc.ErrorCode == -2147417846)
                    {
                        return true;
                    }
                    // otherwise application object is invalid
                    ReleaseExcelApplication();
                    return false;
                }
                catch (AccessViolationException)
                {
                    //release safely because there is a chance to get this exception again
                    ReleaseExcelApplicationSafely();
                    return false;
                }
                catch (Exception)
                {
                    //ExceptionHandler.LogException(ex);
                    // otherwise application object is invalid
                    ReleaseExcelApplication();
                    return false;
                }
            }
        }

        public bool Equals(IApplication obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Application application = (Application)obj;
            return _excelApplication.Equals(application._excelApplication);
        }

        public IRange Intersect(IRange range1, IRange range2)
        {
            using (new EnUsCultureInvoker())
            {
                MSExcel.Range resultRange = _excelApplication.Intersect((MSExcel.Range)range1.RangeObject,
                                                                        (MSExcel.Range)range2.RangeObject);
                return _entityResolver.ResolveRange(resultRange);
            }
        }

        public IRange Union(IRange range1, IRange range2)
        {
            using (new EnUsCultureInvoker())
            {
                MSExcel.Range resultRange = _excelApplication.Union((MSExcel.Range)range1.RangeObject,
                                                                        (MSExcel.Range)range2.RangeObject);
                return _entityResolver.ResolveRange(resultRange);
            }
        }

        //TODO: docuemnt this... it is not Excel event.. what were it comes from and why EnUs not used?....
        public void OnUnsupportedWorkbookChange(IWorkbook oldWorkbook, IWorkbook newWorkbook)
        {
            UnsupportedWorkbookChangeEventHandler handler = UnsupportedWorkbookChange;
            if (handler != null)
            {
                handler(oldWorkbook, newWorkbook);
            }
        }

        public IFileDialog GetFileDialog(FileDialogType fileDialogType)
        {
            using (new EnUsCultureInvoker())
            {
                return _entityResolver.ResolveFileDialog(_excelApplication.FileDialog[MsoFileDialogTypeToFileDialogTypeConverter.ConvertBack(fileDialogType)]);
            }
        }

        public object GetSaveAsFilename(object initialFilename, object fileFilter, object filterIndex, object title, object buttonText)
        {
            using (new EnUsCultureInvoker())
            {
                return _excelApplication.GetSaveAsFilename(initialFilename, fileFilter, filterIndex, title);
            }
        }

        private IEnumerable<int> GetTopLevelWindowHandles(IApplication application)
        {
            ApplicationVersion appVersion = application.ApplicationVersion;
            if (appVersion != Office.Enums.ApplicationVersion.Office2013)
                return new [] { application.Hwnd } ; // in case office version is less or equal MSO2013, there is only one top-level window within an application instance

            return application.Windows.Select(window => 
            {
                //TODO: The code below should be refactored and moved to IWindow interface once a way to retrieve window handle from window object will be found
                MSExcel.Window excelWindow = (MSExcel.Window)window.WindowObject;
                int hwnd = (int)excelWindow.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, null, excelWindow, null);
                return hwnd;
            }).Where(hwnd=> 
            {
                return OfficeWindowsManager.IsMainWindowHandle((IntPtr)hwnd);
            });

        }

        public void SendKeys(object keys, object wait = null)
        {
            using (new EnUsCultureInvoker())
            {
                _excelApplication.SendKeys(keys, wait);
            }
        }

        public double InchesToPoints(double inches)
        {
            using (new EnUsCultureInvoker())
            {
                return _excelApplication.InchesToPoints(inches);
            }
        }

        public void Run(object macro)
        {
            using (new EnUsCultureInvoker())
            {
                _excelApplication.Run(macro);
            }
        }

        public void OnUndo(string text, string procedure)
        {
            using (new EnUsCultureInvoker())
            {
                _excelApplication.OnUndo(text, procedure);
            }
        }

        public void OnRepeat(string text, string procedure)
        {
            using (new EnUsCultureInvoker())
            {
                _excelApplication.OnRepeat(text, procedure);
            }
        }

		public string ActivePrinter
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelApplication.ActivePrinter;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelApplication.ActivePrinter = value;
				}
			}
		}

        #endregion

        #region Events

        public event NewWorkbookEventHandler NewWorkbook
        {
            add { _applicationEvents.NewWorkbook += value; }
            remove { _applicationEvents.NewWorkbook -= value; }
        }

        public event SheetSelectionChangeEventHandler SheetSelectionChange
        {
            add { _applicationEvents.SheetSelectionChange += value; }
            remove { _applicationEvents.SheetSelectionChange -= value; }
        }

        public event SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClick
        {
            add { _applicationEvents.SheetBeforeDoubleClick += value; }
            remove { _applicationEvents.SheetBeforeDoubleClick -= value; }
        }

        public event SheetBeforeRightClickEventHandler SheetBeforeRightClick
        {
            add { _applicationEvents.SheetBeforeRightClick += value; }
            remove { _applicationEvents.SheetBeforeRightClick -= value; }
        }

        event SheetActivateEventHandler IApplication.SheetActivate
        {
            add { _applicationEvents.SheetActivate += value; }
            remove { _applicationEvents.SheetActivate -= value; }
        }

        public event SheetDeactivateEventHandler SheetDeactivate
        {
            add { _applicationEvents.SheetDeactivate += value; }
            remove { _applicationEvents.SheetDeactivate -= value; }
        }

        public event SheetCalculateEventHandler SheetCalculate
        {
            add { _applicationEvents.SheetCalculate += value; }
            remove { _applicationEvents.SheetCalculate -= value; }
        }

        public event SheetChangeEventHandler SheetChange
        {
            add { _applicationEvents.SheetChange += value; }
            remove { _applicationEvents.SheetChange -= value; }
        }

        public event WorkbookOpenEventHandler WorkbookOpen
        {
            add { _applicationEvents.WorkbookOpen += value; }
            remove { _applicationEvents.WorkbookOpen -= value; }
        }

        public event WorkbookActivateEventHandler WorkbookActivate
        {
            add { _applicationEvents.WorkbookActivate += value; }
            remove { _applicationEvents.WorkbookActivate -= value; }
        }

        public event WorkbookDeactivateEventHandler WorkbookDeactivate
        {
            add { _applicationEvents.WorkbookDeactivate += value; }
            remove { _applicationEvents.WorkbookDeactivate -= value; }
        }

        public event WorkbookBeforeCloseEventHandler WorkbookBeforeClose
        {
            add { _applicationEvents.WorkbookBeforeClose += value; }
            remove { _applicationEvents.WorkbookBeforeClose -= value; }
        }

        public event WorkbookBeforeSaveEventHandler WorkbookBeforeSave
        {
            add { _applicationEvents.WorkbookBeforeSave += value; }
            remove { _applicationEvents.WorkbookBeforeSave -= value; }
        }

        //was added only in Excel 2010
        //private WorkbookAfterSaveEventHandler _workbookAfterSave;
        /////<inheritdoc/>
        //public event WorkbookAfterSaveEventHandler WorkbookAfterSave
        //{
        //    add
        //    {
        //        SetupEventsConnection();
        //        _workbookAfterSave += value;
        //    }
        //    remove
        //    {
        //        _workbookAfterSave -= value;
        //    }
        //}

        public event WorkbookBeforePrintEventHandler WorkbookBeforePrint
        {
            add { _applicationEvents.WorkbookBeforePrint += value; }
            remove { _applicationEvents.WorkbookBeforePrint -= value; }
        }

        public event WorkbookNewSheetEventHandler WorkbookNewSheet
        {
            add { _applicationEvents.WorkbookNewSheet += value; }
            remove { _applicationEvents.WorkbookNewSheet -= value; }
        }

        public event WorkbookAddinInstallEventHandler WorkbookAddinInstall
        {
            add { _applicationEvents.WorkbookAddinInstall += value; }
            remove { _applicationEvents.WorkbookAddinInstall -= value; }
        }

        public event WorkbookAddinUninstallEventHandler WorkbookAddinUninstall
        {
            add { _applicationEvents.WorkbookAddinUninstall += value; }
            remove { _applicationEvents.WorkbookAddinUninstall -= value; }
        }

        public event WindowResizeEventHandler WindowResize
        {
            add { _applicationEvents.WindowResize += value; }
            remove { _applicationEvents.WindowResize -= value; }
        }

        public event WindowActivateEventHandler WindowActivate
        {
            add { _applicationEvents.WindowActivate += value; }
            remove { _applicationEvents.WindowActivate -= value; }
        }

        ///<inheritdoc/>
        public event WindowDeactivateEventHandler WindowDeactivate
        {
            add { _applicationEvents.WindowDeactivate += value; }
            remove { _applicationEvents.WindowDeactivate -= value; }
        }

        public event WorkbookRowsetCompleteEventHandler WorkbookRowsetComplete
        {
            add { _applicationEvents.WorkbookRowsetComplete += value; }
            remove { _applicationEvents.WorkbookRowsetComplete -= value; }
        }

        public event UnsupportedWorkbookChangeEventHandler UnsupportedWorkbookChange;

        public event SheetPivotTableUpdateEventHandler SheetPivotTableUpdate
        {
            add { _applicationEvents.SheetPivotTableUpdate += value; }
            remove { _applicationEvents.SheetPivotTableUpdate -= value; }
        }

        public event SheetFollowHyperlinkEventHandler SheetFollowHyperlink
        {
            add { _applicationEvents.SheetFollowHyperlink += value; }
            remove { _applicationEvents.SheetFollowHyperlink -= value; }
        }

        #endregion
        
        #region TRDA Compatibility members

        public bool GetIsInEditMode()
        {
            if (!IsStarted || GetIsModalWindowOpened())
            {
                return false;
            }
            EnableEvents = false;
            object missing = Type.Missing;
            const int newMenu = 18;
            bool isInEditMode = false;

            ICommandBar menuBar = CommandBars["Worksheet Menu Bar"];
            ICommandBarControl newDocumentButton = menuBar.FindControl(
                ControlType.ControlButton, //the type of item to look for
                newMenu, //the item to look for
                missing, //the tag property (in this case missing)
                missing, //the visible property (in this case missing)
                true); //we want to look for it recursively
            EnableEvents = true;
            if (newDocumentButton != null)
            {
                if (!newDocumentButton.Enabled && Workbooks.Count != 0)
                {
                    isInEditMode = true;
                }
            }

            //Check if open embedded object in PowerPoint
            if (Workbooks.Any(workbook => workbook.IsInplace))
            {
                isInEditMode = true;
            }

            return isInEditMode;
        }

        public bool GetIsModalWindowOpened()
        {
            return IsStarted && !Office.Windows.NativeWindowsManager.IsWindowEnabled(new IntPtr(Hwnd));
        }

        #endregion

        
    }
}