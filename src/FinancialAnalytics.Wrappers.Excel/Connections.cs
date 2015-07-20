using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Connections : EntitiesCollectionWrapperBase<IConnections, IWorkbookConnection>, IConnections
    {
        protected ExcelEntityResolver EntityResolver { get; private set; }
        private readonly Object _excelConnections;
        private readonly LateBindingInvoker _invoker;

        public Connections(ExcelEntityResolver entityResolver, Object excelConnections)
        {
            if (entityResolver == null)
            {
                throw new ArgumentNullException("entityResolver");
            }
            if (excelConnections == null)
            {
                throw new ArgumentNullException("excelConnections");
            }
            this.EntityResolver = entityResolver;
            _excelConnections = excelConnections;
            _invoker = new LateBindingInvoker(_excelConnections);
            InitializeCollection();
        }

        private void InitializeCollection()
        {
            using(new EnUsCultureInvoker())
            {
                int itemCount = _invoker.InvokeGetPropertyValue<int>("Count");
                for (int i = 1; i <= itemCount; i ++)
                {
                    IWorkbookConnection wbConnection = EntityResolver.ResolveWorkbookConnection(_invoker.NamedInvoke("Item", i));
                    _items.Add(wbConnection);
                }
            }
            
        }

        public IWorkbookConnection this[string itemName]
        {
            get
            {
                using(new EnUsCultureInvoker())
                {
                    Object parentObject = _invoker.InvokeGetPropertyValue<Object>("Parent");
                    return EntityResolver.ResolveWorkbookConnection(parentObject.GetType().InvokeMember("Connections",
                                                               BindingFlags.GetProperty,
                                                               null, parentObject,
                                                               new object[] {itemName}));
                }
            }
        }

        public override bool Equals(IConnections obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Connections connections = (Connections)obj;
            return _excelConnections.Equals(connections._excelConnections);
        }

        public IWorkbookConnection Add(string name, string description, Object connectionString, Object commandText, Object cmdtype)
        {
             using(new EnUsCultureInvoker())
             {
                 IWorkbookConnection wbConnection = EntityResolver.ResolveWorkbookConnection(_invoker.NamedInvoke("Add", 
                                                                                                                                 name, 
                                                                                                                                 description, 
                                                                                                                                 connectionString, 
                                                                                                                                 commandText,
                                                                                                                                 cmdtype
                                                                                                                             ));
                 _items.Add(wbConnection);
                 return wbConnection;
             }
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
                ComObjectsFinalizer.ReleaseComObject(_excelConnections);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        #region IEnumerable Members

        public new System.Collections.IEnumerator GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        #endregion
    }
}
