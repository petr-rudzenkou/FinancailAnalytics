using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class VisualBasicEditor : ExcelEntityWrapper<IVisualBasicEditor>, IVisualBasicEditor
    {
        protected Microsoft.Vbe.Interop.VBE _interopVisualBasicEditor;

        public VisualBasicEditor(ExcelEntityResolver entityResolver, Microsoft.Vbe.Interop.VBE visualBasicEditor)
            :base(entityResolver)
        {
            if (visualBasicEditor == null)
                throw new ArgumentNullException("visualBasicEditor");
            _interopVisualBasicEditor = visualBasicEditor;
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
				ComObjectsFinalizer.ReleaseComObject(_interopVisualBasicEditor);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public IVisualBasicEditorWindow MainWindow
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveVisulaBasicEditorWindow(_interopVisualBasicEditor.MainWindow);
                }
            }
        }

		public override bool Equals(IVisualBasicEditor obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            VisualBasicEditor visualBasicEditorWindow = (VisualBasicEditor)obj;
            return _interopVisualBasicEditor.Equals(visualBasicEditorWindow._interopVisualBasicEditor);
        }
    }
}
