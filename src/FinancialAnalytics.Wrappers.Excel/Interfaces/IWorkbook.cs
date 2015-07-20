using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.EventsRouting;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    [ComVisible(true)]
    public interface IWorkbook : IEntityWrapper<IWorkbook>
    {
        INames Names { get; }
        ICharts Charts { get; }   
        IWorksheets Worksheets { get; }
        ISheets Sheets { get; }
        ISheet ActiveSheet { get; }
        ICustomDocumentProperties CustomDocumentProperties { get; }
        IApplication Application { get; }
        string FullName { get; }
        object WorkbookObject { get; }
        string Path { get; }
        string Name { get; }
        string Title { get; set; }
        bool Saved { get; set; }
        bool IsInplace { get; }
        void Save();
        void Activate();
        void Close(bool saveChanges);
    	void CloseWithoutAlerts(bool saveChanges);
        /// <summary>
        /// Returns a Windows collection that represents all the windows in the specified workbook. Read-only Windows object.
        /// </summary>
        IWindows Windows { get; }
        bool ReadOnly { get; }
        bool Final { get; }
        bool ProtectStructure { get; }
        IRange ActiveCell { get; }
        IPivotCaches PivotCaches();
        IConnections Connections { get; }
        IChart ActiveChart { get; }
        ITheme Theme { get; }
		IChart GetActiveChart();
		ICustomXmlParts CustomXmlParts { get; }
    	FileFormat FileFormat { get; }
    	bool HasVBProject { get; }

    	void SaveAs(Object filename,
    	            Object fileFormat = null,
    	            Object password = null,
    	            Object writeResPassword = null,
    	            Object readOnlyRecommended = null,
    	            Object createBackup = null,
    	            SaveAsAccessMode accessMode = SaveAsAccessMode.NoChange,
    	            Object conflictResolution = null,
    	            Object addToMru = null,
    	            Object textCodepage = null,
    	            Object textVisualLayout = null,
    	            Object local = null);
		
		void SaveAs(string fileName);
        void SaveCopyAs(string fileName);
        void SaveAsTemplate(string fileName);
        void Close(Object saveChanges, Object filename, Object routeWorkbook);

        bool RemovePersonalInformation { get; set; }
        void Protect(object password, object structure, object windows);
        void Unprotect(object password);

        string UncPath { get; }
        string UncFullName { get; }

		object GetColors();
		void SetColors(object colors);

    	void DeleteNumberFormat(string numberFormat);

        /// <summary>
        /// Returns the names of macros available in this workbook.
        /// This method will throw an error of type, System.Runtime.InteropServices.COMException if "Trsut Access to VBA Project Object Model" option is unchecked on Excel option.
        /// This method is not part of real Excel API, not documented in MSDN.
        /// </summary>
        /// <returns></returns>
        //TODO: move to extensions as not part of real COM API - no a wrapper
        //TODO: fix bad design by returning interface of non modifiable collection 
        List<string> GetMacroNames();
    }
}
