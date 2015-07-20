using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;

using ILegend = FinancialAnalytics.Wrappers.Excel.Interfaces.ILegend;
using ILegendEntries = FinancialAnalytics.Wrappers.Excel.Interfaces.ILegendEntries;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Legend : ExcelEntityWrapper<ILegend>, ILegend
    {
        private Microsoft.Office.Interop.Excel.Legend _excelLegend;

        public Legend(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Legend legend)
            : base(entityResolver)
        {
            if (legend == null)
                throw new ArgumentNullException("legend");
            _excelLegend = legend;
        }

        ~Legend()
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
				ComObjectsFinalizer.ReleaseComObject(_excelLegend);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public double Height
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelLegend.Height;
                }
            }

            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelLegend.Height = value;
                }
            }
        }

        public double Left
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelLegend.Left;
                }
            }

            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelLegend.Left = value;
                }
            }
        }

        public LegendPosition Position 
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    //int aa = (int)_excelLegend.GetType().InvokeMember("Position", System.Reflection.BindingFlags.GetProperty, null, _excelLegend, null);
                    //return XlLegendPositionToLegendPositionConverter.Convert((XlLegendPosition)_excelLegend.Position);
                    try
                    {
                        return
                            XlLegendPositionToLegendPositionConverter.Convert(
                                (int)
                                _excelLegend.GetType().InvokeMember("Position",
                                                                    System.Reflection.BindingFlags.GetProperty, null,
                                                                    _excelLegend, null));
                    }
                    catch
                    {
                        return LegendPosition.LegendPositionCustom;
                    }
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    try
                    {
                        _excelLegend.GetType().InvokeMember("Position", System.Reflection.BindingFlags.SetProperty,
                                                            null, _excelLegend,
                                                            new object[]
                                                                {
                                                                    XlLegendPositionToLegendPositionConverter.
                                                                        ConvertBack(value)
                                                                });
                    }
                    catch
                    {
                        
                    }
                    //_excelLegend.Position = XlLegendPositionToLegendPositionConverter.ConvertBack(value);
                }
            }
        }


        public double Top
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelLegend.Top;
                }
            }

            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelLegend.Top = value;
                }
            }
        }

        public double  Width
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelLegend.Width;
                }
            }

            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelLegend.Width = value;
                }
            }
        }

        public Wrappers.Excel.Interfaces.IFont Font
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveFont(_excelLegend.Font);
                }
            }
        }

        public bool IncludeInLayout
        {
            get
            {
                using (new EnUsCultureInvoker())
                {

                    return _excelLegend.IncludeInLayout; 
                }
            }

            set
            {
                using (new EnUsCultureInvoker())
                {

                    _excelLegend.IncludeInLayout = value;
                }
            }            
        }


        public Interfaces.IInterior Interior
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveInterior(_excelLegend.Interior);
                }
            }
        }


		public Interfaces.IChartFormat Format
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartFormat(_excelLegend.Format);
				}
			}
		}


		public ILegendEntries LegendEntries(object index)
		{
			using (new EnUsCultureInvoker())
			{
				Microsoft.Office.Interop.Excel.LegendEntries entries =
					_excelLegend.LegendEntries(index) as Microsoft.Office.Interop.Excel.LegendEntries;

				if (entries == null)
				{
					return null;
				}

				return EntityResolver.ResolveLegendEntries(entries);
			}
		}

        public override bool Equals(ILegend obj)
        {
            using (new EnUsCultureInvoker())
            {
                if (obj == null || GetType() != obj.GetType())
                {
                    return false;
                }
                Legend chartTitle = (Legend)obj;
                return _excelLegend.Equals(chartTitle._excelLegend);
            }
        }


        public IBorder Border
        {
            get 
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelLegend.Border);
                }
            }
        }
    }
}
