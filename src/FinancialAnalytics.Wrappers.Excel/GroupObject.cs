using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;

namespace FinancialAnalytics.Wrappers.Excel
{
	public class GroupObject : ExcelEntityWrapper<IGroupObject>, IGroupObject
	{
		protected Microsoft.Office.Interop.Excel.GroupObject _excelGroupObject;

		public GroupObject(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.GroupObject groupObject)
			: base(entityResolver)
		{
			if (groupObject == null)
				throw new ArgumentNullException("groupObject");
			_excelGroupObject = groupObject;
		}

		#region Disposable pattern

		private bool disposed = false;
		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelGroupObject);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

		public override bool Equals(IGroupObject obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}

			GroupObject groupObject = (GroupObject)obj;
			return _excelGroupObject.Equals(groupObject._excelGroupObject);
		}

		public IInterior Interior
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveInterior(_excelGroupObject.Interior);
				}
			}
		}

		public IBorder Border
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveBorder(_excelGroupObject.Border);
				}
			}
		}

		public IFont Font
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveFont(_excelGroupObject.Font);
				}
			}
		}
	}
}
