using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office.EventsRouting;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IExcelApplicationLoader
    {
        IEnumerable<IApplication> GetWrappedApplications();
        IEnumerable<IApplication> GetWrappedApplications(IEnumerable<int> handlersToSkip);
        IEnumerable<IWorkbook> GetWorkbooks();
        bool IsExcelStarted();
        void DisposeApplications();
        IWorkbook OpenWorkbook(string path, bool ignoreSameNames);
        IWorkbook OpenWorkbook(string path);
        IApplication GetActiveApplication();
        void CloseApplication(IApplication application);
		void CloseApplicationWithoutAlerts(IApplication application);
		void DisposeApplication(int windowHandle);
        //Added for compatibility with old version which used by TRDA
        bool GetIsInEditMode();
        //Added for compatibility with old version which used by TRDA
        bool GetIsModalWindowOpened();
        //Added for compatibility with old version which used by TRDA
        IWorkbook CreateWorkbook();
        int InstainstancesCount { get; }
        IApplication StartApplicationInHiddenMode();
		event EventHandler<ApplicationInitializedEventArgs> ApplicationInitialized;
    }
}
