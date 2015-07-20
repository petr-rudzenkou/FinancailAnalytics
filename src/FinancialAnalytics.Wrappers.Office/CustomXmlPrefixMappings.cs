using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	internal class CustomXmlPrefixMappings : EntityWrapperBase<ICustomXmlPrefixMappings>, ICustomXmlPrefixMappings
	{
		private Microsoft.Office.Core.CustomXMLPrefixMappings _mappings;

		public CustomXmlPrefixMappings(EntityResolverBase entityResolver, Microsoft.Office.Core.CustomXMLPrefixMappings mappings)
			: base(entityResolver)
		{
			if (mappings == null)
			{
				throw new ArgumentNullException("mappings");
			}
			_mappings = mappings;
		}

		public void AddNamespace(string prefix, string uri)
		{
			_mappings.AddNamespace(prefix, uri);
		}

		public override bool Equals(ICustomXmlPrefixMappings obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			CustomXmlPrefixMappings mappings = (CustomXmlPrefixMappings)obj;
			return _mappings.Equals(mappings._mappings);
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
				ComObjectsFinalizer.ReleaseComObject(_mappings);
				_mappings = null;
				disposed = true;
			}
			base.Dispose(disposing);
		}
	}
}
