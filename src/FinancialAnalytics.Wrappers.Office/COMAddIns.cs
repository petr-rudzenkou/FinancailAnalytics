using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
    public class COMAddIns : EntityWrapperBase<ICOMAddIns>, ICOMAddIns
    {
        protected Microsoft.Office.Core.COMAddIns _COMAddIns;

        public COMAddIns(EntityResolverBase entityResolver, Microsoft.Office.Core.COMAddIns COMAddIns)
            : base(entityResolver)
        {
            if (COMAddIns == null)
                throw new ArgumentNullException("COMAddIns");
            _COMAddIns = COMAddIns;
        }

        public ICOMAddIn Item(ref object objectIndex)
        {
            return _entityResolver.ResolveCOMAddIn(_COMAddIns.Item(ref objectIndex));
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
				ComObjectsFinalizer.ReleaseComObject(_COMAddIns);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public override bool Equals(ICOMAddIns obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            COMAddIns COMAddIns = (COMAddIns)obj;
            return _COMAddIns.Equals(COMAddIns._COMAddIns);
        }
    }
}
