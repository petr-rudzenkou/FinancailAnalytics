using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Utils;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.EventsRouting;
using FinancialAnalytics.Wrappers.Office.ExceptionHandling;
using FinancialAnalytics.Wrappers.Office.Windows;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace FinancialAnalytics.Wrappers.Excel
{

    public class ExcelApplicationLoader : IExcelApplicationLoader
    {
		private IList<Tuple<IApplication, int, uint>> _applications;
        private readonly OfficeVersion _currentVersion;
        private readonly ExcelProcessManager _excelProcessManager;
    	private readonly bool _includeHiddenApplications;
		private ReaderWriterLockSlim _lockerSlim;

		public ExcelApplicationLoader(OfficeVersion officeVersion, bool includeHiddenApplications)
		{
			_applications = new List<Tuple<IApplication, int, uint>>();
			_currentVersion = officeVersion;
			_excelProcessManager = new ExcelProcessManager();
			_includeHiddenApplications = includeHiddenApplications;
			_lockerSlim = new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);
		}

		public ExcelApplicationLoader(OfficeVersion officeVersion)
			: this(officeVersion, false)
		{
		}

        ~ExcelApplicationLoader()
        {
            try
            {
                DisposeApplications();
            }
            catch (Exception)
            {
            }
        }

        public void DisposeApplications()
        {
            DisposeUnmanagedPart();
            GC.SuppressFinalize(this);
        }

        private void CloseApplicationInternal(IApplication application, bool displayAlerts)
        {
			using (LockHelper.WriteLock(_lockerSlim))
			{
				if (_applications.Any(x => x.Item1 == application))
				{
					_applications.Remove(_applications.First(x => x.Item1 == application));
				}
			}
			if (application.IsApplicationValid())
			{
				IWorkbooks workbooks = application.Workbooks;
				if (displayAlerts)
				{
					workbooks.CloseAll();
				}
				else
				{
					workbooks.CloseAllWithoutAlerts();
				}
				workbooks.Dispose();
				application.Quit();
				application.Dispose();
			}
        }

        public void CloseApplication(IApplication application)
        {
            CloseApplicationInternal(application, true);
        }

		public void CloseApplicationWithoutAlerts(IApplication application)
		{
			CloseApplicationInternal(application, false);
		}

		public void DisposeApplication(int windowHandle)
		{
			if (windowHandle <= 0)
			{
				return;
			}
		    try
            {
				using (LockHelper.WriteLock(_lockerSlim))
				{
					var currentItem = _applications.FirstOrDefault(app => app.Item2 == windowHandle);
					IApplication application = null;
					if (currentItem != null)
					{
						application = currentItem.Item1;
					}
					
					if (application != null)
					{
						_applications.Remove(_applications.First(app => app.Item2 == windowHandle));
						application.Dispose();
						GC.Collect();
						GC.WaitForPendingFinalizers();
					}
				}
            }
		    catch (Exception exception)
		    {
                bool rethrow = ExceptionHandler.HandleException(exception);
                if (rethrow)
                    throw;
		    }
		}
		
        public IEnumerable<IApplication> GetWrappedApplications()
        {
            return GetWrappedApplications(new List<int>());
        }

        public IEnumerable<IApplication> GetWrappedApplications(IEnumerable<int> handlersToSkip)
        {
			using (LockHelper.WriteLock(_lockerSlim))
			{
				List<IApplication> invalApplications = _applications
					.Where(x => handlersToSkip.Contains(x.Item2) || !x.Item1.IsApplicationValid()).Select(x => x.Item1).ToList();

				foreach (IApplication invalidApplication in invalApplications)
				{
					_applications.Remove(_applications.First(x => x.Item1 == invalidApplication));
					invalidApplication.Dispose();
				}

				IDictionary<Microsoft.Office.Interop.Excel.Application, uint> nativeApplications =
					_excelProcessManager.GetApplications(_includeHiddenApplications);

				foreach (KeyValuePair<MSExcel.Application, uint> nativeApplicationItem in nativeApplications)
				{
					Microsoft.Office.Interop.Excel.Application nativeApplication = nativeApplicationItem.Key;
					if (handlersToSkip.Contains(nativeApplication.Hwnd) || _applications.Any(x => x.Item2 == nativeApplication.Hwnd) || _applications.Any(x => (x.Item3 != 0) && (x.Item3 == nativeApplicationItem.Value)))
					{
						ComObjectsFinalizer.ReleaseComObject(nativeApplication);
						continue;
					}
					RegisterApplication(new Application(nativeApplication), nativeApplicationItem.Value);
				}
			}
			return _applications.Select(x => x.Item1).ToList();
        }

        public IEnumerable<IWorkbook> GetWorkbooks()
        {
            List<IWorkbook> workbooks = new List<IWorkbook>();
            foreach (IApplication application in GetWrappedApplications())
            {	
				IWorkbooks applicationWorkbooks = application.Workbooks;
				foreach (IWorkbook workbook in applicationWorkbooks)
				{
					ISheet activeSheet = workbook.ActiveSheet;
					if (activeSheet != null)
					{
						activeSheet.Dispose();
						if(IsValidWorkbookName(workbook.FullName))
						{
							workbooks.Add(workbook);
							continue;
						}
					}
					workbook.Dispose();
				}
				applicationWorkbooks.Dispose();
            }
            return workbooks;
        }

        private bool IsValidWorkbookName(string fullName)
        {
            string extension = Path.GetExtension(fullName).ToLower();
            return !((fullName.StartsWith("Chart in") || fullName.StartsWith("Worksheet in"))
                        && (extension.StartsWith(".ppt") || extension.StartsWith(".doc"))
                        );
        }

        public bool IsExcelStarted()
        {
			return _excelProcessManager.GetApplications(_includeHiddenApplications).Any();
        }

        public IWorkbook OpenWorkbook(string fileName)
        {
            return OpenWorkbook(fileName, false);
        }

        /// <summary>
        /// Close existing workbook if it has same name, then open this one.
        /// </summary>
        public IWorkbook OpenWorkbook(string fileName, bool ignoreSameNames)
        {
            IWorkbook result = null;
            foreach (IApplication app in GetWrappedApplications())
            {
                try
                {
	                bool isVisible = app.Visible;
					IWorkbooks workbooks = app.Workbooks;
                    result = workbooks.Open(fileName);
					workbooks.Dispose();
	                //Fix for Excel 2013 - after workbook opening application becomes visible if was hidden
					if (!isVisible)
					{
						app.Visible = false;
					}
					break;
                }
                catch (COMException ex)
                {
                    bool isDuplicateFilename = ex.ErrorCode == -2146827284;
                    if (isDuplicateFilename)
                    {
                        continue;
                    }
                    throw;
                }
            }
            // no any application is started or all applications contains workbook with name Path.GetFileName(fileName)
            if (_applications.Count == 0 || (result == null && ignoreSameNames))
            {
                Application newApplication = new Application(new ApplicationIds(_currentVersion));
				newApplication.AutoLaunch = true;
				newApplication.InitializeApplication();
				RegisterApplication(newApplication);
				IWorkbooks workbooks = newApplication.Workbooks;
				result = workbooks.Open(fileName);
				workbooks.Dispose();
				return result;
            }
            // failed to open the workbook
            if (result == null)
            {
                string targetWorkbookName = System.IO.Path.GetFileName(fileName);
				IApplication firstApplication = _applications.First().Item1;
                IWorkbook workbookToClose = firstApplication.Workbooks.FirstOrDefault(
                        workbook => workbook.Name.Equals(targetWorkbookName, StringComparison.InvariantCultureIgnoreCase));
                if (workbookToClose != null)
                {
                    bool interactive = workbookToClose.Application.Interactive;
                    try
                    {
                        workbookToClose.Application.Interactive = false;
                        workbookToClose.Close(true);
                    }
                    finally
                    {
                        workbookToClose.Application.Interactive = interactive;
                    }
                }
                result = firstApplication.Workbooks.Open(fileName);
            }
            return result;
        }

        public IApplication GetActiveApplication()
        {
            int hwnd = (int)OfficeWindowsManager.FindExcelWindows().FirstOrDefault();

	        var currentItem = _applications.FirstOrDefault(app => app.Item2.Equals(hwnd));
			IApplication wrappedApplication = null;
			if (currentItem != null)
			{
				wrappedApplication = currentItem.Item1;	
			}

            if (wrappedApplication != null)
            {
                return wrappedApplication;
            }
            MSExcel.Application application = _excelProcessManager.GetApplication((IntPtr)hwnd);
            if (application != null)
            {
                wrappedApplication = new Application(application);
				RegisterApplication(wrappedApplication);
            	return wrappedApplication;
            }
            return null;
        }

        private void DisposeUnmanagedPart()
        {
			using (LockHelper.WriteLock(_lockerSlim))
			{
				foreach (IApplication application in _applications.Select(x => x.Item1))
				{
					application.Dispose();
				}
				_applications.Clear();
			}
        }

        public bool GetIsInEditMode()
        {
            IEnumerable<IApplication> excelApplications = GetWrappedApplications();
            return excelApplications.Any(application => application.GetIsInEditMode());
        }

        public bool GetIsModalWindowOpened()
        {
            IEnumerable<IApplication> excelApplications = GetWrappedApplications();
            return excelApplications.Any(application => application.GetIsModalWindowOpened());
        }

        public IWorkbook CreateWorkbook()
        {
            IWorkbook workbook = null;
            if (GetWrappedApplications().Count() == 0)
            {
                IApplication application = new Application(new ApplicationIds(_currentVersion));
				application.AutoLaunch = true;
                application.InitializeApplication();
				RegisterApplication(application);
            	workbook = application.Workbooks.FirstOrDefault();
            }
			return workbook ?? (_applications.First().Item1.Workbooks.Add());
        }

        public int InstainstancesCount
        {
            get
            {
                if (_applications == null)
                    return 0;
                else
                    return _applications.Count();
            }
        }

        public IApplication StartApplicationInHiddenMode()
        {
            if (GetWrappedApplications().Count() == 0)
            {
                IApplication application = new Application(new ApplicationIds(_currentVersion)){ Hidden = true};
				application.AutoLaunch = true;
				application.InitializeApplication();
				application.Visible = false;
            	application.DisplayAlerts = false;
	            RegisterApplication(application);
            }
			return _applications.Select(x => x.Item1).FirstOrDefault();
        }

		/// <summary>
		/// Add application, hwnd and process ID (if passed) to application collection
		/// </summary>
		/// <param name="application"></param>
		/// <param name="processId">0 is added when process Id is unknown for the moment</param>
	    private void RegisterApplication(IApplication application, uint processId = 0)
		{
			using (LockHelper.WriteLock(_lockerSlim))
			{
				int hwnd = application.Hwnd;
				_applications.Add(new Tuple<IApplication, int, uint>(application, hwnd, processId));
				if (ApplicationInitialized != null)
				{
					ApplicationInitialized(this, new ApplicationInitializedEventArgs(hwnd));
				}
			}
		}

		public event EventHandler<ApplicationInitializedEventArgs> ApplicationInitialized;
	}
}
