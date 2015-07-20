using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	public class GradientStops: EntitiesCollectionWrapperBase<IGradientStops, IGradientStop>, IGradientStops
	{
		private readonly Microsoft.Office.Core.GradientStops _officeGradientStops;

		protected EntityResolverBase EntityResolver 
		{ 
			get;
			private set;
		}

		public GradientStops(EntityResolverBase entityResolver, Microsoft.Office.Core.GradientStops officeGradientStops)
		{
			if (officeGradientStops == null)
			{
				throw new ArgumentNullException("officeGradientStops");
			}
			EntityResolver = entityResolver;
			_officeGradientStops = officeGradientStops;

			InitializeCollection();
		}

		private void InitializeCollection()
		{
			for (int i = 1; i <= _officeGradientStops.Count; i++)
			{
				AddItemToCollection(_officeGradientStops[i]);
			}
		}

		private void AddItemToCollection(Microsoft.Office.Core.GradientStop gradientStop)
		{
			_items.Add(EntityResolver.ResolveGradientStop(gradientStop));
		}

		public override bool Equals(IGradientStops obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			GradientStops gradientStops = (GradientStops)obj;
			return this._officeGradientStops.Equals(gradientStops._officeGradientStops);
		}

		public void Insert(int rgb, float position, float transparency = 0f, int index = -1)
		{
			_officeGradientStops.Insert(rgb, position, transparency, index);
		}

		public void Delete(int index)
		{
			_officeGradientStops.Delete(index);
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
				ComObjectsFinalizer.ReleaseComObject(_officeGradientStops);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion
		
	}
}
