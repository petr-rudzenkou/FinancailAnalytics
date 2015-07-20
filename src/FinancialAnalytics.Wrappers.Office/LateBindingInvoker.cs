using System;
using System.Collections.Generic;
using System.Reflection;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Office
{
    //TODO: document when this is thrown
    public class InvokerException : ApplicationException
    {
        public InvokerException() : base() { }
        public InvokerException(string message) : base(message) { }
        public InvokerException(Exception innerException) :
            base("An exception has occurred in invoked method.", innerException) { }
    }

    //TODO: document why it is used and what is fixed, why NetOffice not used in this case - not that late binding degrades start up and runtime
    public class LateBindingInvoker
    {
        
        public object ParentClass { get; private set; }

        public LateBindingInvoker(object parentClass)
        {
            this.ParentClass = parentClass;
        }

        public object NamedInvoke(string methodName, params object[] parameters)
        {
            List<object> paramsConverter = new List<object>(parameters);
            paramsConverter = paramsConverter.ConvertAll<Object>(
                item =>
                    {
                        if (item is IConvertible)
                        {
                            if (item is Enum)
                            {
                                return Convert.ChangeType(item, Enum.GetUnderlyingType(item.GetType()));
                            }
                        }
                        return item;
                    });
            try
            {
                //TODO: investiage get methods info for later invokations for perfromacne reasons
                return ParentClass.GetType()
                    .InvokeMember(methodName, BindingFlags.InvokeMethod | BindingFlags.Default, null, ParentClass, paramsConverter.ToArray());
            }
            catch (Exception e)
            {
                throw new InvokerException(e.InnerException);
            }
        }

		public object InvokeGetPropertyValue(string propertyName, params object[] parameters)
        {
            try
            {
                //TODO: investiage get methods info for later invokations for perfromacne reasons
				return ParentClass.GetType().InvokeMember(propertyName, BindingFlags.GetProperty | BindingFlags.Default, null, ParentClass, parameters);
            }
            catch (Exception e)
            {
                throw new InvokerException(e.InnerException);
            }
        }

        public T InvokeGetPropertyValue<T>(string propertyName)
        {
            try
            {
                return (T)ParentClass.GetType().InvokeMember(propertyName, BindingFlags.GetProperty | BindingFlags.Default, null, ParentClass, null);
            }
            catch (Exception e)
            {
                throw new InvokerException(e.InnerException);
            }
        }

        public void InvokeSetPropertyValue(string propertyName, params object[] values)
        {
            try
            {
				ParentClass.GetType().InvokeMember(propertyName, BindingFlags.SetProperty | BindingFlags.Default, null, ParentClass, values);
            }
            catch (Exception e)
            {
                throw new InvokerException(e.InnerException);
            }
        }

        public static void ReleaseParamsArray(params object[] paramsArray)
        {
            if (null != paramsArray)
            {
                int parramArrayCount = paramsArray.Length;
                for (int i = 0; i < parramArrayCount; i++)
                    ReleaseParam(paramsArray[i]);
            }
        }

        public static void ReleaseParam(object param)
        {
            try
            {
                if (null != param)
                {
                    if (param.GetType().IsSubclassOf(typeof(EntityWrapperBase<>)))
                    {
                        ((IDisposable)param).Dispose();
                    }

                    else if (param is MarshalByRefObject)
                        ComObjectsFinalizer.ReleaseComObject(param);
                }
            }
            catch (Exception throwedException)
            {
                throw new InvokerException(throwedException);
            }
        }
    }

}
