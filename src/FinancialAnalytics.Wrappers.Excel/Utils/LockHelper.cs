using System;
using System.Threading;

namespace FinancialAnalytics.Wrappers.Excel.Utils
{
	internal class LockHelper : IDisposable
	{
		private Action _exitAction;

		private LockHelper(Action exitAction)
		{
			_exitAction = exitAction;
		}

		public static LockHelper ReadLock(ReaderWriterLockSlim readerWriterLockSlim)
		{
			readerWriterLockSlim.EnterReadLock();
			return new LockHelper(readerWriterLockSlim.ExitReadLock);
		}

		public static LockHelper UpgradeableReadLock(ReaderWriterLockSlim readerWriterLockSlim)
		{
			readerWriterLockSlim.EnterUpgradeableReadLock();
			return new LockHelper(readerWriterLockSlim.ExitUpgradeableReadLock);
		}

		public static LockHelper WriteLock(ReaderWriterLockSlim readerWriterLockSlim)
		{
			readerWriterLockSlim.EnterWriteLock();
			return new LockHelper(readerWriterLockSlim.ExitWriteLock);
		}

		public void Dispose()
		{
			if (_exitAction != null)
			{
				_exitAction();
			}
		}
	}
}
