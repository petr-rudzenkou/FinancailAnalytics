using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class PivotCache : ExcelEntityWrapper<IPivotCache>, IPivotCache
    {
        private readonly Microsoft.Office.Interop.Excel.PivotCache _excelPivotCache;
        private readonly LateBindingInvoker _invoker;

        public PivotCache(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.PivotCache pivotCache)
            : base(entityResolver)
        {
            if (pivotCache == null)
            {
                throw new ArgumentNullException("pivotCache");
            }
            _excelPivotCache = pivotCache;
            _invoker = new LateBindingInvoker(_excelPivotCache);
        }

        #region IPivotCache Members

        public bool OLAP
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotCache.OLAP;
                }
            }
        }

        public Object PivotCacheObject
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotCache;
                }
            }
        }

        public object CommandText 
        { 
            get { using (new EnUsCultureInvoker())
            {
                return _excelPivotCache.CommandText;
            }}
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotCache.CommandText = value;
                }
            } 
        }

        public Object Connection
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotCache.Connection;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotCache.Connection = value;
                }
            }
        }

        public bool MaintainConnection
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotCache.MaintainConnection;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotCache.MaintainConnection = value;
                }
            }
        }

        public CmdType CommandType
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlCmdTypeToCmdTypeConverter.Convert(_excelPivotCache.CommandType);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotCache.CommandType = XlCmdTypeToCmdTypeConverter.ConvertBack(value);
                }
            }
        }

        public bool IsConnected
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotCache.IsConnected;
                }
            }
        }

        public IPivotTable CreatePivotTable(
        Object tableDestination,
        Object tableName,
        Object readData,
        Object defaultVersion)
        {
            using (new EnUsCultureInvoker())
            {
                return EntityResolver.ResolvePivotTable(_invoker.NamedInvoke("CreatePivotTable", tableDestination,
                                                                                                        tableName,
                                                                                                        readData,
                                                                                                        defaultVersion
                                                                    ) as Microsoft.Office.Interop.Excel.PivotTable);
            }
        }

        #endregion

        public override bool Equals(IPivotCache obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            PivotCache pivotCache = (PivotCache)obj;
            return _excelPivotCache.Equals(pivotCache._excelPivotCache);
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
                ComObjectsFinalizer.ReleaseComObject(_excelPivotCache);
				_disposed = true;
			}
			base.Dispose(disposing);
		}
		
    }
}
