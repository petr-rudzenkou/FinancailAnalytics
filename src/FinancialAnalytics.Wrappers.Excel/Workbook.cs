using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Collections.Generic;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Vbe.Interop;
using FinancialAnalytics.Wrappers.Office.Interfaces;
using MSExcel = Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.EventsRouting;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Utils;
using FinancialAnalytics.Wrappers.Office;
using ICharts = FinancialAnalytics.Wrappers.Excel.Interfaces.ICharts;
using INames = FinancialAnalytics.Wrappers.Excel.Interfaces.INames;
using IPivotCaches = FinancialAnalytics.Wrappers.Excel.Interfaces.IPivotCaches;
using IRange = FinancialAnalytics.Wrappers.Excel.Interfaces.IRange;
using IWindows = FinancialAnalytics.Wrappers.Excel.Interfaces.IWindows;
using IWorksheets = FinancialAnalytics.Wrappers.Excel.Interfaces.IWorksheets;

namespace FinancialAnalytics.Wrappers.Excel
{
    public class Workbook : ExcelEntityWrapper<IWorkbook>, IWorkbook
    {
        private Microsoft.Office.Interop.Excel._Workbook _excelWorkbook;
        private LateBindingInvoker _invoker;

        public Workbook(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel._Workbook workbook)
            : base(entityResolver)
        {
            if (workbook == null)
                throw new ArgumentNullException("workbook");
            _excelWorkbook = workbook;
            _invoker = new LateBindingInvoker(_excelWorkbook);
			//this.Application.WorkbookBeforeClose += WorkbookCloseInternal;
			//this.Application.WorkbookBeforeSave += WorkbookSaveInternal;
        }
       

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
                //this.Application.WorkbookBeforeClose -= WorkbookCloseInternal;
                //this.Application.WorkbookBeforeSave -= WorkbookSaveInternal;
                //Here we must dispose unmanaged resources and LOH objects
                //lock (_locker)
                //{
                    ComObjectsFinalizer.ReleaseComObject(_excelWorkbook);
					_excelWorkbook = null;
				//}
				_invoker = null;
                _entityResolver = null;
                
                disposed = true;
            }
            
           
            base.Dispose(disposing);
        }
        #endregion

        ~Workbook()
        {
            try
            {
                Dispose(false);
            }
            catch (Exception)
            {
            }
        }
        

        public bool ReadOnly
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorkbook.ReadOnly;
                }
            }
        }

        public string FullName
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorkbook.FullName;
                }
            }
        }

        public string Path
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorkbook.Path;
                }
            }
        }

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorkbook.Name;
                }
            }
        }
		
        public string Title
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorkbook.Title;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelWorkbook.Title = value;
                }
            }
        }

        public bool Saved
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorkbook.Saved;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelWorkbook.Saved = value;
                }
            }
        }

        public bool IsInplace
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorkbook.IsInplace;
                }
            }
        }

        public ISheets Sheets
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveSheets(_excelWorkbook.Sheets);
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

        public object WorkbookObject
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorkbook;
                }
            }
        }

        public ICharts Charts
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveCharts(_excelWorkbook.Charts);
                }
            }
        }

	    public ICustomXmlParts CustomXmlParts
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					Microsoft.Office.Core.CustomXMLParts customXmlParts = _invoker.InvokeGetPropertyValue<Microsoft.Office.Core.CustomXMLParts>("CustomXMLParts");
					return _entityResolver.ResolveCustomXmlParts(customXmlParts);
				}
			}
	    }

        public IConnections Connections
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveConnections(_invoker.InvokeGetPropertyValue("Connections"));
                }
            }
        }

        public IWorksheets Worksheets
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveWorksheets(_excelWorkbook.Worksheets);
                }
            }
        }

        public ISheet ActiveSheet
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return GetActiveSheet();
                }
            }
        }

        public IChart ActiveChart
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    ISheet activeSheet = GetActiveSheet();
					if (GetSheetType(activeSheet) == SheetType.Chart)
					{
                        return (Chart) activeSheet;
                    }
                    return null;
                }
            }
        }

		public IChart GetActiveChart()
		{
			using (new EnUsCultureInvoker())
			{
				if (_excelWorkbook.ActiveChart != null)
				{
					return EntityResolver.ResolveChart(_excelWorkbook.ActiveChart);
				}
				else
				{
					return null;
				}
			}
		}
		
		public INames Names
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveNames(_excelWorkbook.Names);
                }
            }
        }

        public ICustomDocumentProperties CustomDocumentProperties
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveCustomDocumentProperties(_excelWorkbook.CustomDocumentProperties);
                }
            }
        }

        public bool Final
        {
            get
            {
                try
                {
                    using (new EnUsCultureInvoker())
                    {
                        return (bool)_excelWorkbook.GetType().InvokeMember("Final",
                                System.Reflection.BindingFlags.GetProperty, null, _excelWorkbook, new object[] { });
                    }
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

        public bool ProtectStructure
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorkbook.ProtectStructure;
                }
            }
        }

        public IRange ActiveCell
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    if (_excelWorkbook.Windows.Count > 0 && _excelWorkbook.Windows[1].ActiveCell != null)
                    {
                        return EntityResolver.ResolveRange(_excelWorkbook.Windows[1].ActiveCell);
                    }
                    return null;
                }
            }
        }

        public bool RemovePersonalInformation
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorkbook.RemovePersonalInformation;
                }
            }

            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelWorkbook.RemovePersonalInformation = value;
                }
            }
        }

        protected virtual ISheet GetActiveSheet()
        {
            using (new EnUsCultureInvoker())
            {
                SheetType sheetType = GetSheetType(_excelWorkbook.ActiveSheet);
                ISheet activeSheet = null;
                switch (sheetType)
                {
                    case SheetType.Chart:
                        activeSheet = EntityResolver.ResolveChart(_excelWorkbook.ActiveSheet as Microsoft.Office.Interop.Excel.Chart);
                        break;
                    case SheetType.Worksheet:
                        activeSheet = EntityResolver.ResolveWorksheet(_excelWorkbook.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet);
                        break;
                }
                return activeSheet;
            }
        }

    	public FileFormat FileFormat
    	{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlFileFormatToFileFormatConverter.Convert(_excelWorkbook.FileFormat);
				}
			}
    	}

    	public bool HasVBProject
    	{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return CheckVbaProject();
				}
			}
    	}

		private bool CheckVbaProject()
		{
			//Commenting below line as HasVBProject property not available in application object
			//if (Application.Version >= 12) return _excelWorkbook.HasVBProject;
			
			var vbProject = _excelWorkbook.VBProject;
			foreach (VBComponent vbComponent in vbProject.VBComponents)
			{
				if (vbComponent.CodeModule.CountOfLines > 0)
					return true;
			}
			return false;
		}

		public void SaveAs(string fileName)
		{
			SaveAs(fileName,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				SaveAsAccessMode.NoChange,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing); 
		}

        public void SaveCopyAs(string fileName)
		{
            using (new EnUsCultureInvoker())
            {
                _excelWorkbook.SaveCopyAs(fileName);
            }
		}

		public void SaveAsTemplate(string fileName)
		{
			SaveAs(fileName,
				FileFormat.Template,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				true,
				SaveAsAccessMode.NoChange,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing,
				Type.Missing); 
		}

		public void SaveAs(Object filename,
			Object fileFormat,
			Object password,
			Object writeResPassword,
			Object readOnlyRecommended,
			Object createBackup,
			SaveAsAccessMode accessMode,
			Object conflictResolution,
			Object addToMru,
			Object textCodepage,
			Object textVisualLayout,
			Object local)
		{
			using (new EnUsCultureInvoker())
			{
				if (fileFormat != null && fileFormat is FileFormat)
				{
					fileFormat = XlFileFormatToFileFormatConverter.ConvertBack((FileFormat) fileFormat);
				}
				_excelWorkbook._SaveAs(
					filename,
					fileFormat,
					password,
					writeResPassword,
					readOnlyRecommended,
					createBackup,
					XlSaveAsAccessModeToSaveAsAccessModeConverter.ConvertBack(accessMode),
					conflictResolution,
					addToMru,
					textCodepage,
					textVisualLayout
					);
			}
		}

    	public void Close (
	        Object saveChanges,
            Object filename,
            Object routeWorkbook
        )
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorkbook.Close(saveChanges, filename, routeWorkbook);
            }
        }

        public IPivotCaches PivotCaches()
        {
            using (new EnUsCultureInvoker())
            {
               return EntityResolver.ResolvePivotCaches(_excelWorkbook.PivotCaches());
            }
        }

        public void Save()
        {
            using (new EnUsCultureInvoker())
            {
                //we need to prevent unexpected pop-ups during workbook saving
                IApplication excelApplication = EntityResolver.ResolveApplication();
                bool alerts = excelApplication.DisplayAlerts;
                excelApplication.DisplayAlerts = false;
                _excelWorkbook.Save();
                excelApplication.DisplayAlerts = alerts;
            }
        }

        public void Activate()
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorkbook.Activate();
            }
        }

        public void Close(bool saveChanges)
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorkbook.Close(saveChanges, Type.Missing, Type.Missing);
                EntityResolver.ResolveApplication().InitializeApplication();
            }
        }

		public void CloseWithoutAlerts(bool saveChanges)
		{
			using (new EnUsCultureInvoker())
			{
				IApplication excelApplication = EntityResolver.ResolveApplication();
				bool alerts = excelApplication.DisplayAlerts;
				excelApplication.DisplayAlerts = false;
				try
				{
					_excelWorkbook.Close(saveChanges, Type.Missing, Type.Missing);
				}
				finally
				{
					excelApplication.DisplayAlerts = alerts;
					EntityResolver.ResolveApplication().InitializeApplication();
				}
			}
		}

        public override bool Equals(IWorkbook obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Workbook workbook = (Workbook)obj;
            return _excelWorkbook.Equals(workbook._excelWorkbook);
        }

        public IWindows Windows
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveWindows(_excelWorkbook.Windows);
                }
            }
        }

        private static SheetType GetSheetType(object sheet)
        {
            using (new EnUsCultureInvoker())
            {
                SheetType sheetType = SheetType.Other;
                if (sheet is Microsoft.Office.Interop.Excel.Worksheet)
                {
                    sheetType = SheetType.Worksheet;
                }
                else if (sheet is Microsoft.Office.Interop.Excel.Chart)
                {
                    sheetType = SheetType.Chart;
                }
                return sheetType;
            }
        }

        public void Protect(object password, object structure, object windows)
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorkbook.Protect(password, structure, windows);
            }
        }

        public void Unprotect(object password)
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorkbook.Unprotect(password);
            }
        }

        /// <summary>
        /// Returns path to Workbook with replaced mapped drives with network path
        /// (e.g. 'z:' is mapped drive to '\\SomeCompName\SharedDir', for workbook 'z:\Book1.xls' it returns '\\SomeCompName\SharedDir')
        /// </summary>
         public string UncPath
        {
            get { return LocalPathToUncConverter.Convert(Path); }
        }

        /// <summary>
        /// Returns full name of Workbook with replaced mapped drives with network path
        /// (e.g. 'z:' is mapped drive to '\\SomeCompName\SharedDir', for workbook 'z:\Book1.xls' it returns '\\SomeCompName\SharedDir\Book1.xls')
        /// </summary>
        public string UncFullName
        {
            get { return LocalPathToUncConverter.Convert(FullName); }
        }

		public object GetColors()
		{
            using (new EnUsCultureInvoker())
            {
				return _excelWorkbook.get_Colors(Type.Missing);
			}
		}

		public void SetColors(object colors)
		{
			using (new EnUsCultureInvoker())
			{
				_excelWorkbook.set_Colors(Type.Missing, colors);
			}
		}

        public ITheme Theme
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveTheme(_excelWorkbook.Theme);
                }
            }
        }

		public void DeleteNumberFormat(string numberFormat)
		{
			using (new EnUsCultureInvoker())
			{
				_excelWorkbook.DeleteNumberFormat(numberFormat);
			}
		}

        ///<inheritdoc/>
        public List<string> GetMacroNames()
        {
            using (new EnUsCultureInvoker())
            {
               var macroNameList = new List<string>();
                VBProject project = _excelWorkbook.VBProject;
                var procedureType = Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc;
                foreach (var component in project.VBComponents)
                {
                    var vbComponent = component as VBComponent;
                    if (vbComponent != null)
                    {
                        var componentCode = vbComponent.CodeModule;
                        int componentCodeLines = componentCode.CountOfLines;
                        int line = 1;
                        while (line < componentCodeLines)
                        {
                            string procedureName = componentCode.get_ProcOfLine(line, out procedureType);
                            if (!string.IsNullOrEmpty(procedureName))
                            {
                                macroNameList.Add(procedureName);

                                int procedureLines = componentCode.get_ProcCountLines(procedureName, procedureType);
                                line += procedureLines;
                            }
                            else
                            {
                                line += 1;//When line empty move to next TRMO-7973
                            }
                        }
                    }
                }

                return macroNameList;
            }
        }
    }
}
