using System;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Converters;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Shape : ExcelEntityWrapper<IShape>, IShape
    {
        protected Microsoft.Office.Interop.Excel.Shape _excelShape;

        public Shape(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Shape shape)
            : base(entityResolver)
        {
            if (shape == null)
                throw new ArgumentNullException("shape");
            _excelShape = shape;
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
				ComObjectsFinalizer.ReleaseComObject(_excelShape);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

		public int Id
		{	
			get
			{
				using (new EnUsCultureInvoker())
				{	
					return _excelShape.ID; 
				}
			}
		}

		public IApplication Application
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveApplication();
				}
			}
		}		

		public float Height
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelShape.Height;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelShape.Height = value;
                }
            }
        }

        public float Width
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelShape.Width;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelShape.Width = value;
                }
            }
        }

        public ShapeType Type
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return MsoShapeTypeToShapeTypeConverter.Convert(_excelShape.Type);
                }
            }
        }

	    public PlacementType Placement
	    {
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlPlacementToPlacementConverter.Convert(_excelShape.Placement);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelShape.Placement = XlPlacementToPlacementConverter.ConvertBack(value);
				}
			}
	    }

	    public ILineFormat Line
	    {
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveLineFormat(_excelShape.Line);
				}
			}
	    }

        public IFillFormat Fill
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveFillFormat(_excelShape.Fill);
                }
            }
        }

    	public string Name
    	{
    		get
    		{
    			using (new EnUsCultureInvoker())
    			{
    				return _excelShape.Name;
    			}
    		}
    	}

        public void IncrementTop(int value)
        {
            using (new EnUsCultureInvoker())
            {
                _excelShape.IncrementTop(value);
            }
        }

        public void Delete()
        {
            using (new EnUsCultureInvoker())
            {
                _excelShape.Delete();
            }
        }

		public void Cut()
		{
			using (new EnUsCultureInvoker())
			{
				_excelShape.Cut();
			}
		}

		public void Select(bool replace)
		{
			using (new EnUsCultureInvoker())
			{
				_excelShape.Select(MsoTriStateToBoolConverter.ConvertBack(replace));
			}
		}

        public override bool Equals(IShape obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Shape chartTitle = (Shape)obj;
            return _excelShape.Equals(chartTitle._excelShape);
        }
    }
}
