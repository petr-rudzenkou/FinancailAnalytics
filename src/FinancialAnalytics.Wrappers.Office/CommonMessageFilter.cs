using System;

namespace FinancialAnalytics.Wrappers.Office
{
    [Obsolete("Please use StaComCrossThreadInvoker (the same functionality, but more clear how to use)")]
    public class CommonMessageFilter : IMessageFilter, IDisposable
    {
        private IMessageFilter _oldFilter;
        private bool _disposed;
		private uint _maximumTotalWaitTime;

        private const int DEFAULT_RETRY_TIMEOUT = 1;
        public const int CANCEL = -1;

        /// <summary>
        /// Creates and registers this filter.
        /// </summary>
        public CommonMessageFilter()
			: this(uint.MaxValue)
        {
        }

		/// <summary>
		/// Creates and registers this filter.
		/// </summary>
		/// <param name="maximumTotalWaitTime">
        /// Number of milliseconds before message filter stops spin waiting call to finish, call canceled and COM exceptions popup.
		/// </param>
		public CommonMessageFilter(uint maximumTotalWaitTime)
		{
			_maximumTotalWaitTime = maximumTotalWaitTime;
			_oldFilter = null;
			//TODO: handle error result
			int hr = NativeMethods.CoRegisterMessageFilter(this, out _oldFilter);
		}

        [Obsolete]
        public static void Register()
        {
            IMessageFilter newFilter = new CommonMessageFilter();
            IMessageFilter oldFilter = null;
            NativeMethods.CoRegisterMessageFilter(newFilter, out oldFilter);
        }


        [Obsolete]
        public static void Revoke()
        {
            IMessageFilter oldFilter = null;
            NativeMethods.CoRegisterMessageFilter(null, out oldFilter);
        }

        /// <summary>
        /// Unregisters common filter and returns back previous one.
        /// </summary>
        public void Dispose()
        {
            if (!_disposed)
            {
                IMessageFilter shouldBeCommonFilter = null;
                NativeMethods.CoRegisterMessageFilter(_oldFilter, out shouldBeCommonFilter);
                //if (shouldBeCommonFilter != this) throw new InvalidOperationException("Other message filter was registered inside common filter without returning back common filter.");
                _disposed = true;
            }
        }

        public int HandleInComingCall(uint dwCallType, IntPtr htaskCaller, uint dwTickCount, InterfaceInfo[] lpInterfaceInfo)
        {
            return 1;
        }

        public int RetryRejectedCall(IntPtr htaskCallee, uint dwTickCount, uint dwRejectType)
        {
			if (dwTickCount > _maximumTotalWaitTime)
			{
			    return CANCEL;
			}
            if (ShouldRetry(dwTickCount))
            {
                return DEFAULT_RETRY_TIMEOUT;    
            }
            return CANCEL;
        }

        /// <summary>
        /// Ovveride to provide custom logic when retry happens. Can be used to show some notification to user.
        /// </summary>
        /// <param name="totalWaitTime">Total time user already waited call to finish in milliseconds</param>
        /// <returns></returns>
        public virtual bool ShouldRetry(uint totalWaitTime)
        {
            return true;
        }

        public int MessagePending(IntPtr htaskCallee, uint dwTickCount, uint dwPendingType)
        {
            return 1;
        }


    }
}
