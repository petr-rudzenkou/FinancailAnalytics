using FinancialAnalytics.Wrappers.Excel.EventsRouting;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    internal class ApplicationEvents
    {
        private bool _selectionEventsDisabled;

        private readonly SheetSelectionChangeEventRouter _sheetSelectionChangeEventRouter;
        private readonly NewWorkbookEventRouter _newWorkbookEventRouter;
        private readonly WindowResizeEventRouter _windowResizeEventRouter;
        private readonly SheetBeforeDoubleClickEventRouter _sheetBeforeDoubleClickEventRouter;
        private readonly SheetBeforeRightClickEventRouter _sheetBeforeRightClickEventRouter;
        private readonly SheetActivateEventRouter _sheetActivateEventRouter;
        private readonly SheetDeactivateEventRouter _sheetDeactivateEventRouter;
        private readonly SheetCalculateEventRouter _sheetCalculateEventRouter;
        private readonly SheetChangeEventRouter _sheetChangeEventRouter;
        private readonly WorkbookOpenEventRouter _workbookOpenEventRouter;
        private readonly WorkbookActivateEventRouter _workbookActivateEventRouter;
        private readonly WorkbookDeactivateEventRouter _workbookDeactivateEventRouter;
        private readonly WorkbookBeforeCloseEventRouter _workbookBeforeCloseEventRouter;
        private readonly WorkbookBeforeSaveEventRouter _workbookBeforeSaveEventRouter;
        private readonly WorkbookBeforePrintEventRouter _workbookBeforePrintEventRouter;
        private readonly WorkbookNewSheetEventRouter _workbookNewSheetEventRouter;
        private readonly WorkbookAddinInstallEventRouter _workbookAddinInstallEventRouter;
        private readonly WorkbookAddinUninstallEventRouter _workbookAddinUninstallEventRouter;
        private readonly WindowActivateEventRouter _windowActivateEventRouter;
        private readonly WindowDeactivateEventRouter _windowDeactivateEventRouter;
        private readonly SheetPivotTableUpdateEventRouter _sheetPivotTableUpdateEventRouter;
        private readonly SheetFollowHyperlinkEventRouter _sheetFollowHyperlinkEventRouter;
        private readonly WorkbookRowsetCompleteEventRouter _workbookRowsetCompleteEventRouter;

        public ApplicationEvents()
        {
            _sheetSelectionChangeEventRouter = new SheetSelectionChangeEventRouter();
            _newWorkbookEventRouter = new NewWorkbookEventRouter();
            _windowResizeEventRouter = new WindowResizeEventRouter();
            _sheetBeforeDoubleClickEventRouter = new SheetBeforeDoubleClickEventRouter();
            _sheetBeforeRightClickEventRouter = new SheetBeforeRightClickEventRouter();
            _sheetActivateEventRouter = new SheetActivateEventRouter();
            _sheetDeactivateEventRouter = new SheetDeactivateEventRouter();
            _sheetCalculateEventRouter = new SheetCalculateEventRouter();
            _sheetChangeEventRouter = new SheetChangeEventRouter();
            _workbookOpenEventRouter = new WorkbookOpenEventRouter();
            _workbookActivateEventRouter = new WorkbookActivateEventRouter();
            _workbookDeactivateEventRouter = new WorkbookDeactivateEventRouter();
            _workbookBeforeCloseEventRouter= new WorkbookBeforeCloseEventRouter();
            _workbookBeforeSaveEventRouter = new WorkbookBeforeSaveEventRouter();
            _workbookBeforePrintEventRouter = new WorkbookBeforePrintEventRouter();
            _workbookNewSheetEventRouter = new WorkbookNewSheetEventRouter();
            _workbookAddinInstallEventRouter = new WorkbookAddinInstallEventRouter();
            _workbookAddinUninstallEventRouter = new WorkbookAddinUninstallEventRouter();
            _windowActivateEventRouter = new WindowActivateEventRouter();
            _windowDeactivateEventRouter = new WindowDeactivateEventRouter();
            _sheetPivotTableUpdateEventRouter = new SheetPivotTableUpdateEventRouter();
            _sheetFollowHyperlinkEventRouter = new SheetFollowHyperlinkEventRouter();
            _workbookRowsetCompleteEventRouter = new WorkbookRowsetCompleteEventRouter();
        }

        public bool SelectionEventsDisabled
        {
            get { return _selectionEventsDisabled; }
            set
            {
                _selectionEventsDisabled = value;
                _workbookDeactivateEventRouter.IsEnabled = !value;
                _workbookActivateEventRouter.IsEnabled = !value;
                _sheetSelectionChangeEventRouter.IsEnabled = !value;
                _sheetActivateEventRouter.IsEnabled = !value;
                _sheetDeactivateEventRouter.IsEnabled = !value;
            }
        }

        public void Attach(object applicationObject, ExcelEntityResolver entityResolver)
        {
            _sheetSelectionChangeEventRouter.AttachRcw(applicationObject, entityResolver);
            _newWorkbookEventRouter.AttachRcw(applicationObject, entityResolver);
            _windowResizeEventRouter.AttachRcw(applicationObject, entityResolver);
            _sheetBeforeDoubleClickEventRouter.AttachRcw(applicationObject, entityResolver);
            _sheetBeforeRightClickEventRouter.AttachRcw(applicationObject, entityResolver);
            _sheetActivateEventRouter.AttachRcw(applicationObject, entityResolver);
            _sheetDeactivateEventRouter.AttachRcw(applicationObject, entityResolver);
            _sheetCalculateEventRouter.AttachRcw(applicationObject, entityResolver);
            _sheetChangeEventRouter.AttachRcw(applicationObject, entityResolver);
            _workbookOpenEventRouter.AttachRcw(applicationObject, entityResolver);
            _workbookActivateEventRouter.AttachRcw(applicationObject, entityResolver);
            _workbookDeactivateEventRouter.AttachRcw(applicationObject, entityResolver);
            _workbookBeforeCloseEventRouter.AttachRcw(applicationObject, entityResolver);
            _workbookBeforeSaveEventRouter.AttachRcw(applicationObject, entityResolver);
            _workbookBeforePrintEventRouter.AttachRcw(applicationObject, entityResolver);
            _workbookNewSheetEventRouter.AttachRcw(applicationObject, entityResolver);
            _workbookAddinInstallEventRouter.AttachRcw(applicationObject, entityResolver);
            _workbookAddinUninstallEventRouter.AttachRcw(applicationObject, entityResolver);
            _windowActivateEventRouter.AttachRcw(applicationObject, entityResolver);
            _windowDeactivateEventRouter.AttachRcw(applicationObject, entityResolver);
            _sheetPivotTableUpdateEventRouter.AttachRcw(applicationObject, entityResolver);
            _sheetFollowHyperlinkEventRouter.AttachRcw(applicationObject, entityResolver);
            _workbookRowsetCompleteEventRouter.AttachRcw(applicationObject, entityResolver);
        }

        public void Deattach()
        {
            _sheetSelectionChangeEventRouter.DeattachRcw();
            _newWorkbookEventRouter.DeattachRcw();
            _windowResizeEventRouter.DeattachRcw();
            _sheetBeforeDoubleClickEventRouter.DeattachRcw();
            _sheetBeforeRightClickEventRouter.DeattachRcw();
            _sheetActivateEventRouter.DeattachRcw();
            _sheetDeactivateEventRouter.DeattachRcw();
            _sheetCalculateEventRouter.DeattachRcw();
            _sheetChangeEventRouter.DeattachRcw();
            _workbookOpenEventRouter.DeattachRcw();
            _workbookActivateEventRouter.DeattachRcw();
            _workbookDeactivateEventRouter.DeattachRcw();
            _workbookBeforeCloseEventRouter.DeattachRcw();
            _workbookBeforeSaveEventRouter.DeattachRcw();
            _workbookBeforePrintEventRouter.DeattachRcw();
            _workbookNewSheetEventRouter.DeattachRcw();
            _workbookAddinInstallEventRouter.DeattachRcw();
            _workbookAddinUninstallEventRouter.DeattachRcw();
            _windowActivateEventRouter.DeattachRcw();
            _windowDeactivateEventRouter.DeattachRcw();
            _sheetPivotTableUpdateEventRouter.DeattachRcw();
            _sheetFollowHyperlinkEventRouter.DeattachRcw();
            _workbookRowsetCompleteEventRouter.DeattachRcw();
        }

        public event SheetSelectionChangeEventHandler SheetSelectionChange
        {
            add { _sheetSelectionChangeEventRouter.Combine(value); }
            remove { _sheetSelectionChangeEventRouter.Remove(value); }
        }

        public event NewWorkbookEventHandler NewWorkbook
        {
            add { _newWorkbookEventRouter.Combine(value); }
            remove { _newWorkbookEventRouter.Remove(value); }
        }

        public event WindowResizeEventHandler WindowResize
        {
            add { _windowResizeEventRouter.Combine(value); }
            remove { _windowResizeEventRouter.Remove(value); }
        }

        public event SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClick
        {
            add { _sheetBeforeDoubleClickEventRouter.Combine(value); }
            remove { _sheetBeforeDoubleClickEventRouter.Remove(value); }
        }

        public event SheetBeforeRightClickEventHandler SheetBeforeRightClick
        {
            add { _sheetBeforeRightClickEventRouter.Combine(value); }
            remove { _sheetBeforeRightClickEventRouter.Remove(value); }
        }

        public event SheetActivateEventHandler SheetActivate
        {
            add { _sheetActivateEventRouter.Combine(value); }
            remove { _sheetActivateEventRouter.Remove(value); }
        }

        public event SheetDeactivateEventHandler SheetDeactivate
        {
            add { _sheetDeactivateEventRouter.Combine(value); }
            remove { _sheetDeactivateEventRouter.Remove(value); }
        }

        public event SheetCalculateEventHandler SheetCalculate
        {
            add { _sheetCalculateEventRouter.Combine(value); }
            remove { _sheetCalculateEventRouter.Remove(value); }
        }

        public event SheetChangeEventHandler SheetChange
        {
            add { _sheetChangeEventRouter.Combine(value); }
            remove { _sheetChangeEventRouter.Remove(value); }
        }

        public event WorkbookOpenEventHandler WorkbookOpen
        {
            add { _workbookOpenEventRouter.Combine(value); }
            remove { _workbookOpenEventRouter.Remove(value); }
        }

        public event WorkbookActivateEventHandler WorkbookActivate
        {
            add { _workbookActivateEventRouter.Combine(value); }
            remove { _workbookActivateEventRouter.Remove(value); }
        }

        public event WorkbookDeactivateEventHandler WorkbookDeactivate
        {
            add { _workbookDeactivateEventRouter.Combine(value); }
            remove { _workbookDeactivateEventRouter.Remove(value); }
        }

        public event WorkbookBeforeCloseEventHandler WorkbookBeforeClose
        {
            add { _workbookBeforeCloseEventRouter.Combine(value); }
            remove { _workbookBeforeCloseEventRouter.Remove(value); }
        }

        public event WorkbookBeforeSaveEventHandler WorkbookBeforeSave
        {
            add { _workbookBeforeSaveEventRouter.Combine(value); }
            remove { _workbookBeforeSaveEventRouter.Remove(value); }
        }

        public event WorkbookBeforePrintEventHandler WorkbookBeforePrint
        {
            add { _workbookBeforePrintEventRouter.Combine(value); }
            remove { _workbookBeforePrintEventRouter.Remove(value); }
        }

        public event WorkbookNewSheetEventHandler WorkbookNewSheet
        {
            add { _workbookNewSheetEventRouter.Combine(value); }
            remove { _workbookNewSheetEventRouter.Remove(value); }
        }

        public event WorkbookAddinInstallEventHandler WorkbookAddinInstall
        {
            add { _workbookAddinInstallEventRouter.Combine(value); }
            remove { _workbookAddinInstallEventRouter.Remove(value); }
        }

        public event WorkbookAddinUninstallEventHandler WorkbookAddinUninstall
        {
            add { _workbookAddinUninstallEventRouter.Combine(value); }
            remove { _workbookAddinUninstallEventRouter.Remove(value); }
        }

        public event WindowActivateEventHandler WindowActivate
        {
            add { _windowActivateEventRouter.Combine(value); }
            remove { _windowActivateEventRouter.Remove(value); }
        }

        public event WindowDeactivateEventHandler WindowDeactivate
        {
            add { _windowDeactivateEventRouter.Combine(value);}
            remove { _windowDeactivateEventRouter.Remove(value); }
        }

        public event SheetFollowHyperlinkEventHandler SheetFollowHyperlink
        {
            add { _sheetFollowHyperlinkEventRouter.Combine(value); }
            remove { _sheetFollowHyperlinkEventRouter.Remove(value); }
        }

        public event SheetPivotTableUpdateEventHandler SheetPivotTableUpdate
        {
            add { _sheetPivotTableUpdateEventRouter.Combine(value); }
            remove { _sheetPivotTableUpdateEventRouter.Remove(value); }
        }

        public event WorkbookRowsetCompleteEventHandler WorkbookRowsetComplete
        {
            add { _workbookRowsetCompleteEventRouter.Combine(value); }
            remove { _workbookRowsetCompleteEventRouter.Remove(value); }
        }
    }
}
