using System;
using System.Runtime.InteropServices;
using System.Windows;
using DryTools.Execution;
using DryTools.Primitives;
using FinancialAnalytics.Core.Notification;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Core
{
    public class ApplicationProvider : IApplicationProvider
    {
        public ApplicationProvider(IMessageBoxService messageBoxService)
        {
            _messageBoxService = messageBoxService;
        }

        public void SetApplication(IApplication application)
        {
            _application = application;
            OnReady();
        }

        public void WhenReady(Action<IApplication> action)
        {
            Await.When(
                _ => _application != null,
                h => Ready += h,
                h => Ready -= h)
                .Handle(_ => InvokeComSafe(action, true));
        }

        public bool IfReady(Action<IApplication> action, bool showNotAvaliableMessage = true)
        {
            if (IsReady())
            {
                return InvokeComSafe(action, showNotAvaliableMessage);
            }

            //Log.Error("Excel application is not ready" + Environment.NewLine + new StackTrace(true));
            return false;
        }

        public bool IsReady()
        {
            return _application != null;
        }

        public IDisposable SupressUpdating()
        {
            IfReady(a => a.ScreenUpdating = false);
            return Disposable.With(() => IfReady(a => a.ScreenUpdating = true));
        }

        public void Dispose()
        {
            if (_application != null)
            {
                _application.Dispose();
                _application = null;
            }
        }

        public T GetIfReady<T>(
            Func<IApplication, T> getValue,
            T defaultValue = default(T),
            bool showNotAvaliableMessage = true)
        {
            var value = defaultValue;
            return IfReady(a => value = getValue(a), showNotAvaliableMessage) ? value : defaultValue;
        }

        public T GetIfRangeSelected<T>(
            Func<IRange, T> getResult,
            T defaultResult = default(T))
        {
            var result = defaultResult;

            IfReady(app =>
            {
                var range = app.Selection as IRange;
                if (range != null)
                    using (range)
                        result = getResult(range);
            });

            return result;
        }

        #region Implementation

        private event EventHandler Ready;

        private bool InvokeComSafe(Action<IApplication> action, bool showNotAvaliableMessage)
        {
            DateTime startTime = DateTime.Now;

            while (true)
            {
                try
                {
                    if (Run.CheckIsOnUI())
                    {
                        action.Invoke(_application);
                    }
                    else
                    {
                        var timeExecuting = DateTime.Now - startTime;
                        if (_maxRetryTime > timeExecuting)
                        {
                            Run.OnUI(() =>
                            {
                                using (new ConfirmingCommonMessageFilter((uint)(_maxRetryTime - timeExecuting).TotalSeconds, _messageBoxService))
                                {
                                    action.Invoke(_application);
                                }
                            });
                        }
                    }

                    return true;
                }
                catch (COMException ex)
                {
                    var errorCode = (uint)ex.ErrorCode;
                    if (errorCode == 0x800AC472 || errorCode == 0x80010001) // VBA_E_IGNORE or RPC_E_CALL_REJECTED message
                    {
                        if (DateTime.Now - startTime > _maxRetryTime)
                        {
                            //Log.Error(ex);
                            if (showNotAvaliableMessage)
                            {
                                _messageBoxService.Show(Resources.NotificationResources.The_application_is_still_busy_message);
                            }
                            return false;
                        }

                        if (showNotAvaliableMessage)
                        {
                            var msgResult = _messageBoxService.Show("The application is busy. Do you want to repeat the latest operation?",
                                                                    null,
                                                                    MessageBoxButton.YesNo);
                            if (msgResult == MessageBoxResult.No)
                            {
                                return false;
                            }
                        }
                    }
                    else
                    {
                        throw new ApplicationException(
                            "Error occurred during executing of delegate, passed to ApplicationProvider. See inner exception for more info.", ex);
                    }
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(
                        "Error inside delegate, passed to ApplicationProvider occurred. See inner exception for more info.", ex);
                }
            }
        }

        private void OnReady()
        {
            if (_application != null)
            {
                var evt = Ready;
                if (evt != null)
                    evt(this, EventArgs.Empty);
            }
        }

        private IApplication _application;
        private readonly IMessageBoxService _messageBoxService;
        private readonly TimeSpan _maxRetryTime = TimeSpan.FromSeconds(5);

        #endregion

        public IApplication Application
        {
            get { return _application; }
        }
    }
}
