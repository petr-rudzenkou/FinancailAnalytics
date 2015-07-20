using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	internal class CustomXmlNode : EntityWrapperBase<ICustomXmlNode>, ICustomXmlNode
	{
		private Microsoft.Office.Core.CustomXMLNode _customXmlNode;

		public CustomXmlNode(EntityResolverBase entityResolver, Microsoft.Office.Core.CustomXMLNode customXmlNode)
			: base(entityResolver)
		{
			if (customXmlNode == null)
			{
				throw new ArgumentNullException("customXmlNode");
			}
			_customXmlNode = customXmlNode;
		}

		public string Xml
		{
			get { return _customXmlNode.XML; }
		}

		public void AppendChildSubtree(string xml)
		{
			_customXmlNode.AppendChildSubtree(xml);
		}

		public ICustomXmlNodes SelectNodes(string xpath)
		{
			Microsoft.Office.Core.CustomXMLNodes customXmlNodes = _customXmlNode.SelectNodes(xpath);
			return EntityResolver.ResolveCustomXmlNodes(customXmlNodes);
		}

		public ICustomXmlNode SelectSingleNode(string xpath)
		{
			Microsoft.Office.Core.CustomXMLNode customXmlNode = _customXmlNode.SelectSingleNode(xpath);
			return EntityResolver.ResolveCustomXmlNode(customXmlNode);
		}

		public void Delete()
		{
			_customXmlNode.Delete();
		}

		public override bool Equals(ICustomXmlNode obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			CustomXmlNode customXmlPart = (CustomXmlNode)obj;
			return _customXmlNode.Equals(customXmlPart._customXmlNode);
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
				ComObjectsFinalizer.ReleaseComObject(_customXmlNode);
				_customXmlNode = null;
				disposed = true;
			}
			base.Dispose(disposing);
		}
	}
}
