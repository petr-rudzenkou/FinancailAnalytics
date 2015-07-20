using System;
using System.Collections;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	public class DocumentProperties : EntitiesCollectionWrapperBase<IDocumentProperties, IDocumentProperty>, IDocumentProperties
	{
		private readonly object _officeDocumentProperties;

		protected EntityResolverBase EntityResolver { get; private set; }
		private readonly LateBindingInvoker _invoker;

		public DocumentProperties(EntityResolverBase entityResolver, object officeDocumentProperties)
		{
			if (entityResolver == null)
			{
				throw new ArgumentNullException("entityResolver");
			}
			if (officeDocumentProperties == null)
			{
				throw new ArgumentNullException("officeDocumentProperties");
			}
			this.EntityResolver = entityResolver;
			_officeDocumentProperties = officeDocumentProperties;
			_invoker = new LateBindingInvoker(_officeDocumentProperties);
			InitializeCollection();
		}

		private void InitializeCollection()
		{
			int count = _invoker.InvokeGetPropertyValue<int>("Count");
			IEnumerator enumerator = (IEnumerator)_invoker.NamedInvoke("GetEnumerator");
			while(enumerator.MoveNext())
			{
				try
				{
					IDocumentProperty documentProperty = EntityResolver.ResolveDocumentProperty(enumerator.Current);
					_items.Add(documentProperty);
				}
				catch
				{
				}
			}
		}

		public override bool Equals(IDocumentProperties obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			DocumentProperties documentProperties = (DocumentProperties)obj;
			return _officeDocumentProperties.Equals(documentProperties._officeDocumentProperties);
		}

		private bool _disposed;

		protected override void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				if (disposing)
				{
					//Here we must dispose managed resources
				}
				//Here we must dispose unmanaged resources and LOH objects
				ComObjectsFinalizer.ReleaseComObject(_officeDocumentProperties);
				_disposed = true;
			}
			base.Dispose(disposing);
		}

		public IDocumentProperty this[object index]
		{
			get { return EntityResolver.ResolveDocumentProperty(_invoker.NamedInvoke("Item", index)); }
		}

		
	}
}
