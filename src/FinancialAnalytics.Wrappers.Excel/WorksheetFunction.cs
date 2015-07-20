using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;

using IWorksheetFunction = FinancialAnalytics.Wrappers.Excel.Interfaces.IWorksheetFunction;

namespace FinancialAnalytics.Wrappers.Excel
{
	public class WorksheetFunction : ExcelEntityWrapper<IWorksheetFunction>, IWorksheetFunction
	{
		private Microsoft.Office.Interop.Excel.WorksheetFunction _excelworksheetFunction;

		public WorksheetFunction(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.WorksheetFunction worksheetFunction)
			: base(entityResolver)
		{
			if (worksheetFunction == null)
				throw new ArgumentNullException("worksheetFunction");
			_excelworksheetFunction = worksheetFunction;
		}

		~WorksheetFunction()
		{
			try
			{
				Dispose(false);
			}
			catch (Exception)
			{
			}
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
				ComObjectsFinalizer.ReleaseComObject(_excelworksheetFunction);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

		public double Max(object range)
		{
			using (new EnUsCultureInvoker())
			{
				return _excelworksheetFunction.Max(range);
			}
		}

		public double Match(object value, object range, object index)
		{
			using (new EnUsCultureInvoker())
			{
				return _excelworksheetFunction.Match(value, range, index);
			}
		}

		public dynamic Index(object range, double value)
		{
			using (new EnUsCultureInvoker())
			{
				return _excelworksheetFunction.Index(range, value);
			}
		}

		public double Min(object range)
		{
			using (new EnUsCultureInvoker())
			{
				return _excelworksheetFunction.Min(range); 
			}
		}

		public override bool Equals(IWorksheetFunction obj)
		{
			using (new EnUsCultureInvoker())
			{
				if (obj == null || GetType() != obj.GetType())
				{
					return false;
				}
				WorksheetFunction currentFunction = (WorksheetFunction)obj;
				return _excelworksheetFunction.Equals(currentFunction._excelworksheetFunction);
			}
		}
	}
}
