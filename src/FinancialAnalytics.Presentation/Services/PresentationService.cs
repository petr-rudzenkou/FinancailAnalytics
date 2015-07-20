using System;
using System.Threading;
using System.Windows.Threading;
using FinancialAnalytics.Presentation.Core;

namespace FinancialAnalytics.Presentation.Services
{
    public class PresentationService : IPresentationService
    {
        private readonly IOfficeWindowManager _officeWindowManager;
        private Dispatcher _dispatcher;
        private readonly ManualResetEventSlim _initializedWaitHandle = new ManualResetEventSlim(false);
        private readonly object _lock = new object();
        private bool _configured;
        private bool _isExternalDispatcher;
        private bool _disposed;

        public PresentationService(IOfficeWindowManager officeWindowManager)
        {
            _officeWindowManager = officeWindowManager;
        }

        ~PresentationService()
        {
            try
            {
                Dispose();
            }
            catch (Exception)
            {
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                Invoke(() =>
                {
                    try
                    {
                        _officeWindowManager.CloseAllWindows();
                    }
                    catch
                    { }
                });

                if (_dispatcher != null)
                {
                    if (!_isExternalDispatcher)
                    {
                        _dispatcher.InvokeShutdown();
                    }
                    _dispatcher = null;
                }
                _disposed = true;
            }
        }

        public void Invoke(Action action)
        {
            Configure();
            _initializedWaitHandle.Wait();
            _dispatcher.Invoke(action);
        }

        public void BeginInvoke(Action action)
        {
            Configure();
            _initializedWaitHandle.Wait();
            _dispatcher.BeginInvoke(action);
        }

        public void InvokeShutdown()
        {
            if (_configured)
            {
                lock (_lock)
                {
                    if (_configured)
                    {
                        _dispatcher.InvokeShutdown();
                        _initializedWaitHandle.Reset();
                        _configured = false;
                    }
                }
            }
        }

        private void StartDispatcher()
        {
            var staThread = new Thread(x =>
            {
                _dispatcher = Dispatcher.CurrentDispatcher;
                _initializedWaitHandle.Set();
                Dispatcher.Run();
            });

            staThread.IsBackground = true;
            staThread.Name = "Presentation service";
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
        }

        private void Configure()
        {
            if (!_configured)
            {
                lock (_lock)
                {
                    if (!_configured)
                    {
                        StartDispatcher();
                        _configured = true;
                    }
                }
            }
        }
    }
}
