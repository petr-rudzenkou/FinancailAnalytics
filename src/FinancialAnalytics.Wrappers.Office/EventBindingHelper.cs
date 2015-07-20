using System;
using System.Reflection;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Office
{
    public static class EventBindingHelper
    {


        /// <summary>
        /// returns array of all event listeners
        /// </summary>
        /// <param name="eventName">name of event</param>
        public static Delegate[] GetEventRecipients<T>(EntityWrapperBase<T> comObject, string eventName)
        {
            if (comObject == null)
            {
                throw new ArgumentNullException("comObject");
            }
            Type thisType = comObject.GetType();

            FieldInfo eventsFieldInfo = thisType.GetField(
                                                "_" + eventName + "Event",
                                                BindingFlags.Instance |
                                                BindingFlags.NonPublic);
            if (eventsFieldInfo != null)
            {
                MulticastDelegate eventDelegate = (MulticastDelegate)eventsFieldInfo.GetValue(comObject);

                if (null != eventDelegate)
                {
                    Delegate[] delegates = eventDelegate.GetInvocationList();
                    return delegates;
                }
                else
                    return new Delegate[0];
            }
            return new Delegate[0];
        }


        /// <summary>
        /// retuns instance has one or more event recipients
        /// </summary>
        public static bool GetHasEventRecipients<T>(EntityWrapperBase<T> comObject)
        {
            if (comObject == null)
            {
                throw new ArgumentNullException("comObject");
            }
            Type thisType = comObject.GetType();

            foreach (EventInfo item in thisType.GetEvents())
            {
                MulticastDelegate eventDelegate = (MulticastDelegate)thisType.GetType().GetField(item.Name,
                                                                        BindingFlags.NonPublic |
                                                                        BindingFlags.Instance).GetValue(comObject);

                if ((null != eventDelegate) && (eventDelegate.GetInvocationList().Length > 0))
                    return false;
            }

            return false;
        }
      
    }
}
