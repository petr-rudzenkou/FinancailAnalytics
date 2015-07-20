using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Vbe.Converters;
using FinancialAnalytics.Wrappers.Vbe.Enums;
using FinancialAnalytics.Wrappers.Vbe.Interfaces;

namespace FinancialAnalytics.Wrappers.Vbe
{
	internal class VBComponents : EntitiesCollectionWrapperBase<IVBComponents, IVBComponent>, IVBComponents
	{
		private Microsoft.Vbe.Interop.VBComponents _vbComponents;

		protected VbeEntityResolver VbeEntityResolver
		{
			get;
			private set;
		}

		public VBComponents(VbeEntityResolver vbeEntityResolver, Microsoft.Vbe.Interop.VBComponents vbComponents)
		{
			if (vbComponents == null)
			{
				throw new ArgumentNullException("VBComponents");
			}
			VbeEntityResolver = vbeEntityResolver;
			_vbComponents = vbComponents;

			InitializeCollection();
		}

		private void InitializeCollection()
		{
			for (int i = 1; i <= _vbComponents.Count; i++)
			{
				AddItemToCollection(_vbComponents.Item(i));
			}
		}

		private void AddItemToCollection(Microsoft.Vbe.Interop.VBComponent vbComponent)
		{
			_items.Add(VbeEntityResolver.ResolveVBComponent(vbComponent));
		}

		public override bool Equals(IVBComponents obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			VBComponents vbComponents = (VBComponents)obj;
			return this._vbComponents.Equals(vbComponents._vbComponents);
		}

		public IVBComponent Add(VbComponentType componentType)
		{
			IVBComponent newComponent = VbeEntityResolver.ResolveVBComponent(_vbComponents.Add(MsoVbComponentTypeToVbComponentTypeConverter.ConvertBack(componentType)));
			_items.Add(newComponent);
			return newComponent;
		}

		public void Remove(IVBComponent vbComponent)
		{
			_vbComponents.Remove((Microsoft.Vbe.Interop.VBComponent)vbComponent.VBComponentObject);
			if (_items.Contains(vbComponent))
			{
				_items.Remove(vbComponent);
			}
			vbComponent.Dispose();
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
				ComObjectsFinalizer.ReleaseComObject(_vbComponents);
				_vbComponents = null;
				disposed = true;
			}
			base.Dispose(disposing);
		}
	}
}
