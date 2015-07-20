using System;
using System.Collections.Generic;
using System.Linq;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office.EventsRouting;

namespace FinancialAnalytics.Wrappers.Excel
{
    public class SingleExcelApplicationLoader : IExcelApplicationLoader
    {
        private readonly IApplication _application;

        public SingleExcelApplicationLoader(IApplication application)
        {
            _application = application;
			if (ApplicationInitialized != null)
			{
				ApplicationInitialized(this, new ApplicationInitializedEventArgs(_application.Hwnd));
			}
        }

        public IEnumerable<IApplication> GetWrappedApplications()
        {
            return new List<IApplication> {_application};
        }

        public IEnumerable<IWorkbook> GetWorkbooks()
        {
            return _application.Workbooks;
        }

        public void DisposeApplications()
        {
            _application.Dispose();
        }

        public bool IsExcelStarted()
        {
            return _application.IsStarted;
        }

        public IWorkbook OpenWorkbook(string path)
        {
            return OpenWorkbook(path, false);
        }

        public IWorkbook OpenWorkbook(string path, bool ignoreSameNames)
        {
            return _application.Workbooks.Open(path);
        }

        public IApplication GetActiveApplication()
        {
            return _application;
        }

        public void CloseApplication(IApplication application)
        {
            _application.Workbooks.CloseAll();
            _application.Quit();
        }

		public void CloseApplicationWithoutAlerts(IApplication application)
		{
			_application.Workbooks.CloseAllWithoutAlerts();
			_application.Quit();
		}

        public IEnumerable<IApplication> GetWrappedApplications(IEnumerable<int> handlersToSkip)
        {
            IList<IApplication> applications = new List<IApplication>();
            if (handlersToSkip.Contains(_application.WindowHandle))
            {
                applications.Add(_application);
            }
            return applications;
        }

		public void DisposeApplication(int windowHandle)
		{
			if (windowHandle == _application.Hwnd)
			{
				_application.Dispose();
			}
		}

        public bool GetIsInEditMode()
        {
            return _application.GetIsInEditMode();
        }

        public bool GetIsModalWindowOpened()
        {
            return _application.GetIsModalWindowOpened();
        }

        public IWorkbook CreateWorkbook()
        {
            return _application.Workbooks.Add();
        }

        public int InstainstancesCount
        {
            get { return _application == null ? 0 : 1; }
        }


        public IApplication StartApplicationInHiddenMode()
        {
            throw new System.NotImplementedException();
        }

		public event EventHandler<ApplicationInitializedEventArgs> ApplicationInitialized;

	}
}
