using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	internal class CustomXmlNodes : EntityWrapperBase<ICustomXmlNodes>, ICustomXmlNodes
	{
		private Microsoft.Office.Core.CustomXMLNodes _customXmlNodes;


		public CustomXmlNodes(EntityResolverBase entityResolver, Microsoft.Office.Core.CustomXMLNodes customXmlNodes)
			: base(entityResolver)
		{
			if (customXmlNodes == null)
			{
				throw new ArgumentNullException("customXmlNodes");
			}
			_customXmlNodes = customXmlNodes;
		}

		public int Count
		{
			get { return _customXmlNodes.Count; }
		}

		public ICustomXmlNode this[int index]
		{
			get
			{
				Microsoft.Office.Core.CustomXMLNode customXmlNode = _customXmlNodes[index];
				return EntityResolver.ResolveCustomXmlNode(customXmlNode);
			}
		}

		public override bool Equals(ICustomXmlNodes obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			CustomXmlNodes customXmlPart = (CustomXmlNodes)obj;
			return this._customXmlNodes.Equals(customXmlPart._customXmlNodes);
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
				ComObjectsFinalizer.ReleaseComObject(_customXmlNodes);
				_customXmlNodes = null;
				disposed = true;
			}
			base.Dispose(disposing);
		}
	}
}
