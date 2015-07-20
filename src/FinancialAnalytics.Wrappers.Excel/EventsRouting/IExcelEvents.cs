using System.Runtime.InteropServices;
using MSExcel = Microsoft.Office.Interop.Excel;
using MSOffice = Microsoft.Office.Core;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
    GuidAttribute("00024413-0000-0000-C000-000000000046")]
    public interface IExcelEvents
    {
        [DispId(0x0000061D)]
        void OnNewWorkbook(MSExcel._Workbook oWB);
        [DispId(0x00000616)]
        void OnSheetSelectionChange([MarshalAs(UnmanagedType.IDispatch)] object oSheet, MSExcel.Range oTarget);
        [DispId(0x00000617)]
        void OnSheetBeforeDoubleClick([MarshalAs(UnmanagedType.IDispatch)] object oSheet, MSExcel.Range oTarget, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x00000618)]
        void OnSheetBeforeRightClick([MarshalAs(UnmanagedType.IDispatch)] object oSheet, MSExcel.Range oTarget, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x00000619)]
        void OnSheetActivate([MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        [DispId(0x0000061A)]
        void OnSheetDeactivate([MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        [DispId(0x0000061B)]
        void OnSheetCalculate([MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        [DispId(0x0000061C)]
        void OnSheetChange([MarshalAs(UnmanagedType.IDispatch)] object oSheet, MSExcel.Range oTarget);
        [DispId(0x0000061F)]
        void OnWorkbookOpen(MSExcel._Workbook oWB);
        [DispId(0x00000620)]
        void OnWorkbookActivate(MSExcel._Workbook oWB);
        [DispId(0x00000621)]
        void OnWorkbookDeactivate(MSExcel._Workbook oWB);
        [DispId(0x00000622)]
        void OnWorkbookBeforeClose(MSExcel._Workbook oWB, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x00000623)]
        void OnWorkbookBeforeSave(MSExcel._Workbook oWB, [MarshalAs(UnmanagedType.VariantBool)]  bool SaveUI, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x00000624)]
        void OnWorkbookBeforePrint(MSExcel._Workbook oWB, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x00000625)]
        void OnWorkbookNewSheet(MSExcel._Workbook oWB, [MarshalAs(UnmanagedType.IDispatch)] object oSheet);
        [DispId(0x00000626)]
        void OnWorkbookAddinInstall(MSExcel._Workbook oWB);
        [DispId(0x00000627)]
        void OnWorkbookAddinUninstall(MSExcel._Workbook oWB);
        [DispId(0x00000612)]
        void OnWindowResize(MSExcel._Workbook oWB, MSExcel.Window oWn);
        [DispId(0x00000614)]
        void OnWindowActivate(MSExcel._Workbook oWB, MSExcel.Window oWn);
        [DispId(0x00000615)]
        void OnWindowDeactivate(MSExcel._Workbook oWB, MSExcel.Window oWn);
        [DispId(0x0000073E)]
        void OnSheetFollowHyperlink([MarshalAs(UnmanagedType.IDispatch)] object oSheet, MSExcel.Hyperlink oTarget);
        [DispId(0x0000086D)]
        void OnSheetPivotTableUpdate([MarshalAs(UnmanagedType.IDispatch)] object oSheet, MSExcel.PivotTable oTarget);
        [DispId(0x00000870)]
        void OnWorkbookPivotTableCloseConnection(MSExcel._Workbook oWB, MSExcel.PivotTable oTarget);
        [DispId(0x00000871)]
        void OnWorkbookPivotTableOpenConnection(MSExcel._Workbook oWB, MSExcel.PivotTable oTarget);
        [DispId(0x000008F1)]
        void OnWorkbookSync(MSExcel._Workbook oWB, MSOffice.MsoSyncEventType SyncType);
        [DispId(0x000008F2)]
        void OnWorkbookBeforeXmlImport(MSExcel._Workbook oWB, MSExcel.XmlMap oMap, string sUrl, [MarshalAs(UnmanagedType.VariantBool)]  bool IsRefresh, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x000008F3)]
        void OnWorkbookAfterXmlImport(MSExcel._Workbook oWB, MSExcel.XmlMap oMap, [MarshalAs(UnmanagedType.VariantBool)]  bool IsRefresh, MSExcel.XlXmlImportResult Result);
        [DispId(0x000008F4)]
        void OnWorkbookBeforeXmlExport(MSExcel._Workbook oWB, MSExcel.XmlMap oMap, string sUrl, [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
        [DispId(0x000008F5)]
        void OnWorkbookAfterXmlExport(MSExcel._Workbook oWB, MSExcel.XmlMap oMap, string sUrl, MSExcel.XlXmlExportResult Result);
        [DispId(0x00000A33)]
        void OnWorkbookRowsetComplete(MSExcel._Workbook oWB, string sDesciption, string sSheet, [MarshalAs(UnmanagedType.VariantBool)]  bool Success);
        [DispId(0x00000A34)]
        void OnAfterCalculate();
    }

}
