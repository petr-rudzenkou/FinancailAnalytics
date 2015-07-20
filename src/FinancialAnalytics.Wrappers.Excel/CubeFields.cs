using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class CubeFields : EntitiesCollectionWrapperBase<ICubeFields, ICubeField>, ICubeFields
    {
        private readonly MSExcel.CubeFields _excelCubeFields;
        protected ExcelEntityResolver Entityresolver { get; private set; }

        public CubeFields(ExcelEntityResolver entityResolver, MSExcel.CubeFields excelCubeFields)
        {
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            if (excelCubeFields == null)
                throw new ArgumentNullException("excelCubeFields");
            Entityresolver = entityResolver;
            _excelCubeFields = excelCubeFields;
            InitializeCollection();
        }

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelCubeFields.Count; i ++)
                {
                    AddItemToCollection(_excelCubeFields[i]);
                }
            }
        }

        private void AddItemToCollection(MSExcel.CubeField item)
        {
            _items.Add(
                Entityresolver.ResolveCubeField(item)
                );
        }


        private bool _disposed = false;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelCubeFields);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        public override bool Equals(ICubeFields obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            CubeFields cubeFields = (CubeFields)obj;
            return _excelCubeFields.Equals(cubeFields._excelCubeFields);
        }

        public ICubeField this[Object index]
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return Entityresolver.ResolveCubeField(_excelCubeFields[index]);
                }
            }
        }
    }
}
