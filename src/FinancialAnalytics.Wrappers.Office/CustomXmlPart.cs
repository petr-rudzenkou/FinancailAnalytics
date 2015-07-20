using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	internal class CustomXmlPart : EntityWrapperBase<ICustomXmlPart>, ICustomXmlPart
	{
		private Microsoft.Office.Core.CustomXMLPart _customXmlPart;

		public CustomXmlPart(EntityResolverBase entityResolver, Microsoft.Office.Core.CustomXMLPart customXmlPart)
			: base(entityResolver)
		{
			if (customXmlPart == null)
			{
				throw new ArgumentNullException("customXmlPart");
			}
			_customXmlPart = customXmlPart;
		}

		public string Id
		{
			get { return _customXmlPart.Id; }
		}

		public string Xml
		{
			get { return _customXmlPart.XML; }
		}

		public ICustomXmlPrefixMappings NamespaceManager
		{
			get { return EntityResolver.ResolveCustomXmlPrefixMappings(_customXmlPart.NamespaceManager); }
		}

		public ICustomXmlNodes SelectNodes(string xpath)
		{
			Microsoft.Office.Core.CustomXMLNodes customXmlNodes = _customXmlPart.SelectNodes(xpath);
			return EntityResolver.ResolveCustomXmlNodes(customXmlNodes);
		}

		public ICustomXmlNode SelectSingleNode(string xpath)
		{
			Microsoft.Office.Core.CustomXMLNode customXmlNode = _customXmlPart.SelectSingleNode(xpath);
			return EntityResolver.ResolveCustomXmlNode(customXmlNode);
		}

		public override bool Equals(ICustomXmlPart obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			CustomXmlPart customXmlPart = (CustomXmlPart)obj;
			return this._customXmlPart.Equals(customXmlPart._customXmlPart);
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
				ComObjectsFinalizer.ReleaseComObject(_customXmlPart);
				_customXmlPart = null;
				disposed = true;
			}
			base.Dispose(disposing);
		}
	}
}
