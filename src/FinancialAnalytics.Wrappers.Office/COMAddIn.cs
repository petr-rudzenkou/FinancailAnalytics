using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
    public class COMAddIn : EntityWrapperBase<ICOMAddIn>, ICOMAddIn
    {
        protected Microsoft.Office.Core.COMAddIn _COMAddIn;

        public COMAddIn(EntityResolverBase entityResolver, Microsoft.Office.Core.COMAddIn COMAddIn)
            : base(entityResolver)
        {
            if (COMAddIn == null)
                throw new ArgumentNullException("COMAddIn");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _COMAddIn = COMAddIn;
        }

        public object Object
        {
            get { return _COMAddIn.Object; }
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
				ComObjectsFinalizer.ReleaseComObject(_COMAddIn);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public override bool Equals(ICOMAddIn obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            COMAddIn COMAddIn = (COMAddIn)obj;
            return _COMAddIn.Equals(COMAddIn._COMAddIn);
        }
        
        public bool Connect
        {
            get { return _COMAddIn.Connect; }
        }
    }
}
