using System;
using System.Collections.Generic;
using System.Windows.Forms;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel
{
    /// <summary>
    /// TRMO-6706 Eikon Excel: VBA project from the closed workbook is present in VBA editor
    /// Store each workbook's wrapper that was created in event handlers and runs dispose method on it.
    /// We do it after 250 interval to make sure that disposing is being done after all events handlers were called.
    /// </summary>
    internal class WorkbooksCollector
    {
        private readonly Timer _timer;
        private readonly List<IWorkbook> _workbooks = new List<IWorkbook>();

        public WorkbooksCollector()
        {
            _timer = new Timer();
            _timer.Interval = 250;
            _timer.Tick += FullDispose;

        }

        public void Add(IWorkbook workbook)
        {
            lock (_workbooks)
            {
                _workbooks.Add(workbook);
            }
        }

        public void Start()
        {
            lock (_workbooks)
            {
                _timer.Start();
            }
        }

        public void Stop()
        {
            lock (_workbooks)
            {
                _timer.Stop();
            }
        }

        private void FullDispose(object sender, EventArgs args)
        {
            lock (_workbooks)
            {
                if (_workbooks.Count != 0)
                {
                    foreach (var workbook in _workbooks)
                    {
                        workbook.Dispose();
                    }
                    _workbooks.Clear();
                }
                _timer.Stop();
            }
        }
    }
}
