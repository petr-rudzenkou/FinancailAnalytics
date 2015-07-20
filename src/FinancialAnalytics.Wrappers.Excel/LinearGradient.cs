using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
	public class LinearGradient : ExcelEntityWrapper<ILinearGradient>, ILinearGradient
	{

		private readonly object _excelGradient;
		private readonly LateBindingInvoker _invoker;

		public LinearGradient(ExcelEntityResolver entityResolver, object excelGradient)
			: base(entityResolver)
		{
			if (entityResolver == null)
			{
				throw new ArgumentNullException("entityResolver");
			}

			if (excelGradient == null)
			{
				throw new ArgumentNullException("excelGradient");
			}

			_excelGradient = excelGradient;
			_invoker = new LateBindingInvoker(_excelGradient);
		}

		#region Properties

		private const string DegreePropertyName = "Degree";
		public double Degree
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _invoker.InvokeGetPropertyValue<double>(DegreePropertyName);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_invoker.InvokeSetPropertyValue(DegreePropertyName, value);
				}
			}
		}

		private const string ColorStopsPropertyName = "ColorStops";
		public IColorStops ColorStops
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveColorStops(_invoker.InvokeGetPropertyValue(ColorStopsPropertyName));
				}
			}
		}

		#endregion Properties

		public override bool Equals(ILinearGradient obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			LinearGradient other = (LinearGradient)obj;
			return _excelGradient.Equals(other._excelGradient);
		}

		#region Disposable pattern

		private bool _disposed;
		protected override void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelGradient);
				_disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

	}
}
