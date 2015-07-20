using System;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using IBorder = FinancialAnalytics.Wrappers.Excel.Interfaces.IBorder;
using IChartFillFormat = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartFillFormat;
using IDataLabel = FinancialAnalytics.Wrappers.Excel.Interfaces.IDataLabel;
using IInterior = FinancialAnalytics.Wrappers.Excel.Interfaces.IInterior;
using FinancialAnalytics.Wrappers.Excel.Converters;

namespace FinancialAnalytics.Wrappers.Excel
{

    internal class DataLabel : ExcelEntityWrapper<IDataLabel>, IDataLabel
    {
        protected Microsoft.Office.Interop.Excel.DataLabel _excelDataLabel;

        public DataLabel(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.DataLabel dataLabel)
            : base(entityResolver)
        {
            if (dataLabel == null)
                throw new ArgumentNullException("dataLabel");
            _excelDataLabel = dataLabel;
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
                ComObjectsFinalizer.ReleaseComObject(_excelDataLabel);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion
        
        public double Left
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelDataLabel.Left;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.Left = value;
                }
            } 
        }

        public double Top
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelDataLabel.Top;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.Top = value;
                }
            }
        }

        public object Separator
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelDataLabel.Separator;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.Separator = value;
                }
            }
        }

        public object Type
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelDataLabel.Type;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.Type = value;
                }
            }
        }

        public bool ShowBubbleSize
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelDataLabel.ShowBubbleSize;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.ShowBubbleSize = value;
                }
            }
        }

        public bool ShowCategoryName
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelDataLabel.ShowCategoryName;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.ShowCategoryName = value;
                }
            }
        }

        public bool ShowLegendKey
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelDataLabel.ShowLegendKey;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.ShowLegendKey = value;
                }
            }
        }

        public bool ShowPercentage
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelDataLabel.ShowPercentage;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.ShowPercentage = value;
                }
            }
        }

        public bool ShowSeriesName
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelDataLabel.ShowSeriesName;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.ShowSeriesName = value;
                }
            }
        }

        public bool ShowValue
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelDataLabel.ShowValue;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.ShowValue = value;
                }
            }
        }

        public DataLabelPosition Position
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlDataLabelPositionToDataLabelPosition.Convert(_excelDataLabel.Position);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.Position = XlDataLabelPositionToDataLabelPosition.ConvertBack(value);
                }
            }        
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelDataLabel.Border);
                }
            } 
        }
        
        public IInterior Interior
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveInterior(_excelDataLabel.Interior);
                }
            }
        }

        public IChartFillFormat Fill
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChartFillFormat(_excelDataLabel.Fill);
                }
            }
        }

        public string Text
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelDataLabel.Text;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelDataLabel.Text = value;
                }
            }
        }

        public object Select()
        {
            using (new EnUsCultureInvoker())
            {
                return _excelDataLabel.Select();
            }
        }

        /// <summary>
        /// Returns a Font object that represents the font of the specified object.
        /// </summary>
        public Wrappers.Excel.Interfaces.IFont Font
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveFont(_excelDataLabel.Font);
                }
            }
        }

        public override bool Equals(IDataLabel obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            DataLabel dataLabel = (DataLabel)obj;
            return _excelDataLabel.Equals(dataLabel);
        }



		public Interfaces.IChartFormat Format
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartFormat(_excelDataLabel.Format);
				}
			}
		}
	}
}
