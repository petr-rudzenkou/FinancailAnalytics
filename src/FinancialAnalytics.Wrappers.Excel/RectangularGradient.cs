using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
	public class RectangularGradient : ExcelEntityWrapper<IRectangularGradient>, IRectangularGradient
	{

		private readonly object _excelGradient;
		private readonly LateBindingInvoker _invoker;

		public RectangularGradient(ExcelEntityResolver entityResolver, object excelGradient)
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

		private const string RectangleTopPropertyName = "RectangleTop";
		public double RectangleTop
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _invoker.InvokeGetPropertyValue<double>(RectangleTopPropertyName);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_invoker.InvokeSetPropertyValue(RectangleTopPropertyName, value);
				}
			}
		}

		private const string RectangleBottomPropertyName = "RectangleBottom";
		public double RectangleBottom
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _invoker.InvokeGetPropertyValue<double>(RectangleBottomPropertyName);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_invoker.InvokeSetPropertyValue(RectangleBottomPropertyName, value);
				}
			}
		}

		private const string RectangleLeftPropertyName = "RectangleLeft";
		public double RectangleLeft
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _invoker.InvokeGetPropertyValue<double>(RectangleLeftPropertyName);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_invoker.InvokeSetPropertyValue(RectangleLeftPropertyName, value);
				}
			}
		}

		private const string RectangleRightPropertyName = "RectangleRight";
		public double RectangleRight
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _invoker.InvokeGetPropertyValue<double>(RectangleRightPropertyName);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_invoker.InvokeSetPropertyValue(RectangleRightPropertyName, value);
				}
			}
		}

		#endregion Properties

		public override bool Equals(IRectangularGradient obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			RectangularGradient other = (RectangularGradient)obj;
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
