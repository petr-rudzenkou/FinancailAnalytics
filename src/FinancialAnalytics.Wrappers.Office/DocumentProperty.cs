using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	public class DocumentProperty : EntityWrapperBase<IDocumentProperty>, IDocumentProperty
	{
		private object _officeDocumentProperty;
		private readonly LateBindingInvoker _invoker;

		public DocumentProperty(EntityResolverBase entityResolver, object documentProperty)
			: base(entityResolver)
		{
			if (documentProperty == null)
			{
				throw new ArgumentNullException("documentProperty");
			}
			_officeDocumentProperty = documentProperty;
			_invoker = new LateBindingInvoker(_officeDocumentProperty);
		}

		public string Name 
		{
			get
			{
				return _invoker.InvokeGetPropertyValue<string>("Name");
			}
		}

		public object Value 
		{
			get
			{
				return _invoker.InvokeGetPropertyValue<object>("Value");
			}
		}

		public override bool Equals(IDocumentProperty obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			DocumentProperty documentProperty = (DocumentProperty)obj;
			return this._officeDocumentProperty.Equals(documentProperty._officeDocumentProperty);
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
				ComObjectsFinalizer.ReleaseComObject(_officeDocumentProperty);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion
	}
}
