using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.EventsRouting;
using FinancialAnalytics.Wrappers.Office.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    [ComVisible(true)]
    public interface IApplication : IApplicationCommon
    {
        IRange GetCaller();
        void Calculate();

        IWorkbooks Workbooks { get; }

        IWorkbook ActiveWorkbook { get; }

		ISheet ActiveSheet { get; }

    	IChart ActiveChart { get; }

        IRange ActiveCell { get; }

		string ActivePrinter { get; set; }

		bool AllowHidden { get; set; }

        bool ScreenUpdating { get; set; }

        IVisualBasicEditor VBE { get; }

        IWindows Windows { get; }

        int Hinstance { get; }

        long HinstancePtr { get;  }

        bool EnableEvents { get; set; }

        bool UserControl { get; set; }

        bool DisplayAlerts { get; set; }

        object Selection { get; }

        object StatusBar { get; set; }

        int Hwnd { get; }

        bool IsMultipleSelection { get; }

        WindowState WindowState { get; set; }

        Calculation Calculation { get; set; }

        IWindow ActiveWindow { get; }

        bool Equals(IApplication obj);

        object ApplicationObject { get; }

        IRange Intersect(IRange range1, IRange range2);

        IRange Union(IRange range1, IRange range2);

        bool IsApplicationValid();

        void OnUnsupportedWorkbookChange(IWorkbook oldWorkbook, IWorkbook newWorkbook);

        CultureInfo UICulture { get; }

        object SelectionObject { get; }

        void Disable();

        void Enable();

        //Added for compatibility with old version which used by TRDA
        bool GetIsInEditMode();

        //Added for compatibility with old version which used by TRDA
        bool GetIsModalWindowOpened();

        bool SelectionEventsDisabled { get; set; }

        bool Interactive { get; set; }

        MousePointer Cursor { get; set; }

        ReferenceStyle ReferenceStyle { get; }

        IAutoRecover AutoRecover { get; }

        bool Ready { get; }

        bool DisplayClipboardWindow { get; set; }

        object GetSaveAsFilename(object initialFilename = null, object fileFilter = null, object filterIndex = null,
                                 object title = null, object buttonText = null);

        object Evaluate(object value);

		IEnumerable<int> TopLevelWindowsHandles { get; }

        void SendKeys(object keys, object wait = null);

        double InchesToPoints(double inches);

		void Run(object macro);

        void OnUndo(string text, string procedure);

        void OnRepeat(string text, string procedure);

		IWorksheetFunction WorksheetFunction { get; }

        #region Events

        event NewWorkbookEventHandler NewWorkbook;
        event SheetSelectionChangeEventHandler SheetSelectionChange;
        event SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClick;
        event SheetBeforeRightClickEventHandler SheetBeforeRightClick;
        event SheetActivateEventHandler SheetActivate;
        /// <summary>
        /// Occurs when any sheet is deactivated.
        /// </summary>
        event SheetDeactivateEventHandler SheetDeactivate;
        event SheetCalculateEventHandler SheetCalculate;
        event SheetChangeEventHandler SheetChange;
        event WorkbookOpenEventHandler WorkbookOpen;
        event WorkbookActivateEventHandler WorkbookActivate;
        /// <summary>
        /// Occurs when any open workbook is deactivated.
        /// </summary>
        event WorkbookDeactivateEventHandler WorkbookDeactivate;

        /// <summary>
        /// Occurs immediately before any open workbook closes.
        /// </summary>
        event WorkbookBeforeCloseEventHandler WorkbookBeforeClose;
        /// <summary>
        /// Occurs before any open workbook is saved.
        /// </summary>
        event WorkbookBeforeSaveEventHandler WorkbookBeforeSave;
        event WorkbookBeforePrintEventHandler WorkbookBeforePrint;
        event WorkbookNewSheetEventHandler WorkbookNewSheet;
        event WorkbookAddinInstallEventHandler WorkbookAddinInstall;
        event WorkbookAddinUninstallEventHandler WorkbookAddinUninstall;
        event WindowResizeEventHandler WindowResize;
        event WindowActivateEventHandler WindowActivate;

        /// <summary>
        /// Occurs when any workbook window is deactivated.
        /// </summary>
        event WindowDeactivateEventHandler WindowDeactivate;
        event WorkbookRowsetCompleteEventHandler WorkbookRowsetComplete;
        event UnsupportedWorkbookChangeEventHandler UnsupportedWorkbookChange;
        event SheetPivotTableUpdateEventHandler SheetPivotTableUpdate;
        event SheetFollowHyperlinkEventHandler SheetFollowHyperlink;

        #endregion
    }
}