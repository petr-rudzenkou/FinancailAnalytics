using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    internal abstract class BaseEvent
    {
        private readonly Guid _eventsUid;
        private readonly int _dispId;
        private volatile object _rcw;
        private long _subscribersCount;
        protected readonly object Lock;
        protected volatile ExcelEntityResolver EntityResolver;
        protected Delegate PublicEvent;
        protected Delegate PrivateEvent;
        private volatile bool _isEnabled;

        protected BaseEvent(int dispId, Guid eventsUid)
        {
            _dispId = dispId;
            Lock = new object();
            IsEnabled = true;
            _eventsUid = eventsUid;
        }

        public bool IsEnabled
        {
            get { return _isEnabled; }
            set { _isEnabled = value; }
        }

        public void AttachRcw(object rcw, ExcelEntityResolver entityResolver)
        {
            lock (Lock)
            {
                DeattachRcw();
                _rcw = rcw;
                EntityResolver = entityResolver;
                if (Interlocked.Read(ref _subscribersCount) > 0)
                {
                    ComEventsHelper.Combine(rcw, _eventsUid, _dispId, PrivateEvent);
                }
            }
        }

        public void DeattachRcw()
        {
            lock (Lock)
            {
                object rcw = _rcw;
                if (rcw != null && Interlocked.Read(ref _subscribersCount) > 0)
                {
                    ComEventsHelper.Remove(rcw, _eventsUid, _dispId, PrivateEvent);
                    _rcw = null;
                }
            }
        }

        public void Combine(Delegate value)
        {
            Delegate combined = Delegate.Combine(PublicEvent, value);
            Interlocked.Exchange(ref PublicEvent, combined);
            long subscribersCount = Interlocked.Increment(ref _subscribersCount);
            //If is first subscriber
            if (subscribersCount == 1)
            {
                lock (Lock)
                {
                    object rcw = _rcw;
                    if (rcw != null)
                    {
                        ComEventsHelper.Combine(rcw, _eventsUid, _dispId, PrivateEvent);
                    }
                }
            }
        }

        public void Remove(Delegate value)
        {
            long subscribersCount = Interlocked.Decrement(ref _subscribersCount);
            //If is last subscriber
            if (subscribersCount == 0)
            {
                lock (Lock)
                {
                    object rcw = _rcw;
                    if (rcw != null)
                    {
                        ComEventsHelper.Remove(rcw, _eventsUid, _dispId, PrivateEvent);
                    }
                }
            }
            Delegate combined = Delegate.Remove(PublicEvent, value);
            Interlocked.Exchange(ref PublicEvent, combined);
        }
    }
}
