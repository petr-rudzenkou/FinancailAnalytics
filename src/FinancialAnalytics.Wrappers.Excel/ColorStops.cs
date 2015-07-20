using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
	public class ColorStops : EntitiesCollectionWrapperBase<IColorStops, IColorStop>, IColorStops
	{
		private readonly object _excelColorStops;
		private readonly ExcelEntityResolver _entityResolver;
		private readonly LateBindingInvoker _invoker;

		public ColorStops(ExcelEntityResolver entityResolver, object excelColorStops)
		{
			if (entityResolver == null)
			{
				throw new ArgumentNullException("entityResolver");
			}

			if (excelColorStops == null)
			{
				throw new ArgumentNullException("excelColorStops");
			}

			_entityResolver = entityResolver;
			_excelColorStops = excelColorStops;
			_invoker = new LateBindingInvoker(_excelColorStops);
			InitializeCollection();
		}

		public override bool Equals(IColorStops obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			ColorStops other = (ColorStops)obj;
			return _excelColorStops.Equals(other._excelColorStops);
		}

		#region Disposable pattern

		private bool _disposed;
		protected override void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelColorStops);
				_disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

		private const string CountPropertyName = "Count";
		private const string ItemPropertyName = "Item";
		private const string ClearMethodName = "Clear";
		private const string AddMethodName = "Add";
		private void InitializeCollection()
		{
			using (new EnUsCultureInvoker())
			{
				int itemCount = _invoker.InvokeGetPropertyValue<int>(CountPropertyName);
				for (int i = 1; i <= itemCount; i++)
				{
					IColorStop item = _entityResolver.ResolveColorStop(_invoker.NamedInvoke(ItemPropertyName, i));
					_items.Add(item);
				}
			}
		}

		public void Clear()
		{
			using (new EnUsCultureInvoker())
			{
				_invoker.NamedInvoke(ClearMethodName);
				_items.Clear();
			}
		}

		public IColorStop Add(double position)
		{
			IColorStop item = _entityResolver.ResolveColorStop(_invoker.NamedInvoke(AddMethodName, position));
			_items.Add(item);
			return item;
		}
	}
}
