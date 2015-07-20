using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.ExceptionHandling;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class CustomDocumentProperties : ExcelEntityWrapper<ICustomDocumentProperties>, ICustomDocumentProperties
    {
        protected object _excelCustomDocumentProperties;

        private const string CustomPropertyNameProperty = "Name";
        private const string CustomPropertyValueProperty = "Value";
        private const string CustomPropertyDeleteMethod = "Delete";
        private const string CustomPropertiesAddMethod = "Add";

        public CustomDocumentProperties(ExcelEntityResolver entityResolver, object customDocumentProperties)
            : base(entityResolver)
        {
            if (customDocumentProperties == null)
                throw new ArgumentNullException("customDocumentProperties");
            _excelCustomDocumentProperties = customDocumentProperties;
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
				ComObjectsFinalizer.ReleaseComObject(_excelCustomDocumentProperties);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public void AddProperty(string propertyName, string value)
        {
            using (new EnUsCultureInvoker())
            {
                object[] args = {
                                    propertyName, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                                    value
                                };

                try
                {
                    _excelCustomDocumentProperties.GetType().InvokeMember(CustomPropertiesAddMethod,
                                                                          System.Reflection.BindingFlags.Default |
                                                                          System.Reflection.BindingFlags.InvokeMethod,
                                                                          null, _excelCustomDocumentProperties, args);
                }
                catch (Exception exc)
                {
                    bool rethrow = ExceptionHandler.HandleException(exc);
                    if (rethrow)
                        throw;
                }
            }
        }

        public void DeleteProperty(string propertyName)
        {
            using (new EnUsCultureInvoker())
            {
                foreach (object documentProperty in (System.Collections.IEnumerable) _excelCustomDocumentProperties)
                {

                    try
                    {
                        string name = documentProperty.GetType().InvokeMember(CustomPropertyNameProperty,
                                                                              System.Reflection.BindingFlags.GetProperty |
                                                                              System.Reflection.BindingFlags.Instance,
                                                                              null, documentProperty, new object[] {}).
                            ToString();
                        if (name.Equals(propertyName, StringComparison.InvariantCulture))
                        {
                            documentProperty.GetType().InvokeMember(CustomPropertyDeleteMethod,
                                                                    System.Reflection.BindingFlags.InvokeMethod |
                                                                    System.Reflection.BindingFlags.Instance, null,
                                                                    documentProperty, new object[] {});
                        }
                    }
                    catch (Exception exc)
                    {
                        bool rethrow = ExceptionHandler.HandleException(exc);
                        if (rethrow)
                            throw;
                    }
                }
            }
        }

        public string GetPropertyValue(string propertyName)
        {
            using (new EnUsCultureInvoker())
            {
                foreach (object documentProperty in (System.Collections.IEnumerable) _excelCustomDocumentProperties)
                {
                    try
                    {
                        string name = documentProperty.GetType().InvokeMember(CustomPropertyNameProperty,
                                                                              System.Reflection.BindingFlags.GetProperty |
                                                                              System.Reflection.BindingFlags.Instance,
                                                                              null, documentProperty, new object[] {}).
                            ToString();
                        if (name.Equals(propertyName, StringComparison.InvariantCulture))
                        {
                            string value = documentProperty.GetType().InvokeMember(CustomPropertyValueProperty,
                                                                                   System.Reflection.BindingFlags.
                                                                                       GetProperty |
                                                                                   System.Reflection.BindingFlags.
                                                                                       Instance, null, documentProperty,
                                                                                   new object[] {}).ToString();
                            return value;
                        }
                    }
                    catch (Exception exc)
                    {
                        bool rethrow = ExceptionHandler.HandleException(exc);
                        if (rethrow)
                            throw;
                    }
                }
                return string.Empty;
            }
        }

        public override bool Equals(ICustomDocumentProperties obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            CustomDocumentProperties chartObjects = (CustomDocumentProperties)obj;
            return _excelCustomDocumentProperties.Equals(chartObjects._excelCustomDocumentProperties);
        }
    }
}
