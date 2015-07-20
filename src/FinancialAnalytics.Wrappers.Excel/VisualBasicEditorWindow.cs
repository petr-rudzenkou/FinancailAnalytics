using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class VisualBasicEditorWindow : ExcelEntityWrapper<IVisualBasicEditorWindow>, IVisualBasicEditorWindow
    {
        protected Microsoft.Vbe.Interop.Window _visualBasicEditorWindow;

        public VisualBasicEditorWindow(ExcelEntityResolver entityResolver, Microsoft.Vbe.Interop.Window window)
            : base(entityResolver)
        {
            if (window == null)
                throw new ArgumentNullException("window");
            _visualBasicEditorWindow = window;
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
				//Here we must dispose unmanaged resources and LOH objects
				ComObjectsFinalizer.ReleaseComObject(_visualBasicEditorWindow);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public WindowState WindowState
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return VbeWindowStateToWindowStateConverter.Convert(_visualBasicEditorWindow.WindowState);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _visualBasicEditorWindow.WindowState = VbeWindowStateToWindowStateConverter.ConvertBack(value);
                }
            }
        }

        public bool Visible
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _visualBasicEditorWindow.Visible;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _visualBasicEditorWindow.Visible = value;
                }
            }
        }

		public override bool Equals(IVisualBasicEditorWindow obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            VisualBasicEditorWindow visualBasicEditorWindow = (VisualBasicEditorWindow)obj;
            return _visualBasicEditorWindow.Equals(visualBasicEditorWindow._visualBasicEditorWindow);
        }
    }
}
