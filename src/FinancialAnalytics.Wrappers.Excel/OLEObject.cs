using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Converters;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class OLEObject : ExcelEntityWrapper<IOLEObject>, IOLEObject 
    {
        protected Microsoft.Office.Interop.Excel.OLEObject _excelOLEObj;

        public OLEObject(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.OLEObject oleObj)
            : base(entityResolver)
        {
            if (oleObj == null)
                throw new ArgumentNullException("OLEObject");
            _excelOLEObj = oleObj;
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
                ComObjectsFinalizer.ReleaseComObject(_excelOLEObj);
                disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion

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

        public double Height
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelOLEObj.Height;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelOLEObj.Height = value;
                }
            }
        }

        public double Width
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelOLEObj.Width;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelOLEObj.Width = value;
                }
            }
        }

        public double Left
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelOLEObj.Left;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelOLEObj.Left = value;
                }
            }
        }

        public double Top
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelOLEObj.Top;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelOLEObj.Top = value;
                }
            }
        }

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelOLEObj.Name;
                }
            }
        }

		public string ProgID
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelOLEObj.progID;
				}
			}
		}

		public object Object
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelOLEObj.Object;
				}
			}
		}

		public object Parent
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					Microsoft.Office.Interop.Excel.Worksheet parentWorksheet = _excelOLEObj.Parent as Microsoft.Office.Interop.Excel.Worksheet;
					if (parentWorksheet != null)
					{
						return EntityResolver.ResolveWorksheet(parentWorksheet);
					}
					return null;
				}
			}
		}

        public IOLEObject Activate()
        {
            using (new EnUsCultureInvoker())
            {
				Microsoft.Office.Interop.Excel.OLEObject result = _excelOLEObj.Activate() as Microsoft.Office.Interop.Excel.OLEObject;
	            return EntityResolver.ResolveOLEObject(result);
            }
        }

        public IOLEObject Delete()
        {
            using (new EnUsCultureInvoker())
            {
				Microsoft.Office.Interop.Excel.OLEObject result = _excelOLEObj.Delete() as Microsoft.Office.Interop.Excel.OLEObject;
				return EntityResolver.ResolveOLEObject(result);
            }
        }

        public IOLEObject Select(bool replace)
        {
            using (new EnUsCultureInvoker())
            {
				Microsoft.Office.Interop.Excel.OLEObject result = _excelOLEObj.Select(replace) as Microsoft.Office.Interop.Excel.OLEObject;
				return EntityResolver.ResolveOLEObject(result);
            }
        }

        public void CopyPicture(PictureAppearance appearance, CopyPictureFormat format)
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.XlPictureAppearance xlPictureAppearance =
                    XlPictureAppearanceToPictureAppearanceConverter.ConvertBack(appearance);
                Microsoft.Office.Interop.Excel.XlCopyPictureFormat xlCopyPictureFormat =
                    XlCopyPictureFormatToCopyPictureFormatConverter.ConvertBack(format);
                RepeatedCopyHelper.ExecuteCopyRepeated(() => _excelOLEObj.CopyPicture(xlPictureAppearance, xlCopyPictureFormat));
            }
        }

        public void Copy()
        {
            using (new EnUsCultureInvoker())
            {
                RepeatedCopyHelper.ExecuteCopyRepeated(() => _excelOLEObj.Copy());
            }
        }

        public override bool Equals(IOLEObject obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            OLEObject chartTitle = (OLEObject)obj;
            return _excelOLEObj.Equals(chartTitle._excelOLEObj);
        }

    }
}
