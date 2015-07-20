using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
	public class ColorStop : ExcelEntityWrapper<IColorStop>, IColorStop
	{
		private readonly object _excelColorStop;
		private readonly LateBindingInvoker _invoker;

		public ColorStop(ExcelEntityResolver entityResolver, object excelColorStop)
			: base(entityResolver)
		{
			if (entityResolver == null)
			{
				throw new ArgumentNullException("entityResolver");
			}

			if (excelColorStop == null)
			{
				throw new ArgumentNullException("excelColorStop");
			}

			_excelColorStop = excelColorStop;
			_invoker = new LateBindingInvoker(_excelColorStop);
		}

		#region Properties

		private const string ColorPropertyName = "Color";
		public object Color
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _invoker.InvokeGetPropertyValue(ColorPropertyName);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_invoker.InvokeSetPropertyValue(ColorPropertyName, value);
				}
			}
		}

		private const string ThemeColorPropertyName = "ThemeColor";
		public int ThemeColor
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _invoker.InvokeGetPropertyValue<int>(ThemeColorPropertyName);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_invoker.InvokeSetPropertyValue(ThemeColorPropertyName, value);
				}
			}
		}

		private const string PositionPropertyName = "Position";
		public double Position
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _invoker.InvokeGetPropertyValue<double>(PositionPropertyName);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_invoker.InvokeSetPropertyValue(PositionPropertyName, value);
				}
			}
		}

		#endregion Properties

		public override bool Equals(IColorStop obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			ColorStop other = (ColorStop)obj;
			return _excelColorStop.Equals(other._excelColorStop);
		}

		#region Disposable pattern

		private bool _disposed;
		protected override void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelColorStop);
				_disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

	}
}
