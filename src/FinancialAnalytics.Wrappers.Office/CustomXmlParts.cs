using System;
using System.Collections.Generic;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	internal class CustomXmlParts : EntityWrapperBase<ICustomXmlParts>, ICustomXmlParts
	{
		private Microsoft.Office.Core.CustomXMLParts _customXmlParts;

		public CustomXmlParts(EntityResolverBase entityResolver, Microsoft.Office.Core.CustomXMLParts customXmlParts)
			: base(entityResolver)
		{
			if (customXmlParts == null)
			{
				throw new ArgumentNullException("customXmlParts");
			}
			_customXmlParts = customXmlParts;
		}

		public int Count
		{
			get { return _customXmlParts.Count; }
		}

		public ICustomXmlPart this[object id]
		{
			get
			{
				Microsoft.Office.Core.CustomXMLPart part = _customXmlParts[id];
				return EntityResolver.ResolveCustomXmlPart(part);
			}
		}

		public ICustomXmlPart Add(string xml)
		{
			Microsoft.Office.Core.CustomXMLPart part = _customXmlParts.Add(xml);
			return _entityResolver.ResolveCustomXmlPart(part);
		}

		public ICustomXmlParts SelectByNamespace(string @namespace)
		{
			Microsoft.Office.Core.CustomXMLParts parts = _customXmlParts.SelectByNamespace(@namespace);
			return EntityResolver.ResolveCustomXmlParts(parts);
		}

		public ICustomXmlPart SelectById(string id)
		{
			Microsoft.Office.Core.CustomXMLPart part = _customXmlParts.SelectByID(id);
			return EntityResolver.ResolveCustomXmlPart(part);
		}

		public override bool Equals(ICustomXmlParts obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			CustomXmlParts customXmlParts = (CustomXmlParts)obj;
			return this._customXmlParts.Equals(customXmlParts._customXmlParts);
		}

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
				ComObjectsFinalizer.ReleaseComObject(_customXmlParts);
				_customXmlParts = null;
				disposed = true;
			}
			base.Dispose(disposing);
		}

		public IEnumerator<ICustomXmlPart> GetEnumerator()
		{
			IList<ICustomXmlPart> customXmlParts = new List<ICustomXmlPart>();
			foreach (Microsoft.Office.Core.CustomXMLPart customXmlPart in _customXmlParts)
			{
				customXmlParts.Add(EntityResolver.ResolveCustomXmlPart(customXmlPart));
			}
			return customXmlParts.GetEnumerator();
		}

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return GetEnumerator();
		}
	}
}
