using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Caliburn.Micro;
using DryTools.Execution;
using DryTools.Primitives;
using FinancialAnalytics.Presentation.UI.Window;

namespace FinancialAnalytics.Presentation.Core
{
    public class OfficeWindowManager : WindowManager, IOfficeWindowManager
    {
        public bool? ShowDialog(object rootModel, Action<Window> configureWindow)
        {
            var window = CreateWindow(rootModel, false, null, null);

            window.WindowStartupLocation = ConfigureWindow.DefaultWindowStartupLocation;
            window.ResizeMode = ResizeMode.NoResize;

            if (configureWindow != null)
                configureWindow(window);

            bool? dialogResult = null;
            Run.OnUIAndWait(() => dialogResult = window.ShowDialog());
            return dialogResult;
        }

        public void ShowPopup(object rootModel)
        {
            Run.OnUI(() => ShowPopup(rootModel, null));
        }

        public void ShowWindow(object viewModel, Action<Window> configureWindow = null, Action<NativeMessagesInterceptor> configureInterceptor = null)
        {
            var window = base.CreateWindow(viewModel, false , null, null);

            //issue with closing windows. Don't figure out why does it happen but it fixes the problem.
            window.ResizeMode = ResizeMode.CanResize;
            window.AllowsTransparency = true;

            if (configureWindow != null)
                configureWindow(window);

            Run.OnUIAndWait(() => window.ShowModeless()).ThrowOnError();
        }

        public bool ShowWindowOnlyWhenPreviousIsClosed(string windowId, object viewModel, Action<Window> configureWindow = null)
        {
            Window window;
            if (_openedWindows.TryGetValue(windowId, out window))
            {
                // Paranoid check.
                if (window.IsVisible)
                    return false;

                _openedWindows.Remove(windowId);
            }

            ShowAndRegisterInOpenWindows(windowId, viewModel, configureWindow);
            return true;
        }

        public void ShowWindowAndClosePrevious(string windowId, object viewModel, Action<Window> configureWindow = null)
        {
            CloseWindow(windowId);
            ShowAndRegisterInOpenWindows(windowId, viewModel, configureWindow);
        }

        public bool IsWindowActive(string windowId)
        {
            Window window;
            if (_openedWindows.TryGetValue(windowId, out window))
            {
                return true;
            }
            return false;
        }

        public void CloseAllWindows()
        {
            foreach (var key in _openedWindows.Keys)
            {
                CloseWindow(key);
            }
        }

        public bool TryRegisterWindow(string windowId, Window window)
        {
            if (_openedWindows.ContainsKey(windowId))
            {
                return false;
            }

            _openedWindows.Add(windowId, window);
            window.Closed += (s, e) => _openedWindows.Remove(windowId);

            return true;
        }

        public bool TryActivateWindow(string windowId)
        {
            Window existingWindow;
            _openedWindows.TryGetValue(windowId, out existingWindow);
            if (existingWindow != null && existingWindow.IsVisible)
            {
                existingWindow.Activate();
                return true;
            }

            return false;
        }

        #region WindowManager overrides

        protected override Window EnsureWindow(object model, object view, bool isDialog)
        {
            var windowBase = view as WindowBase;
            var userControl = view as UserControl;
            var commonOfficeStyledViewBase = view as FinancialAnalyticsStyledViewBase;
            ResizeMode resMode = isDialog ? ResizeMode.CanResize : ResizeMode.NoResize;
            if (commonOfficeStyledViewBase != null)
            {
                windowBase = new WindowBase
                {
                    ResizeMode = resMode,
                    Content = view,
                    Width = commonOfficeStyledViewBase.ActualWidth,
                    Height = commonOfficeStyledViewBase.ActualHeight
                    //SizeToContent = SizeToContent.WidthAndHeight
                };
                if (commonOfficeStyledViewBase.HeaderContent != null)
                {
                    //windowBase.HeaderContent = commonOfficeStyledViewBase.HeaderContent;
                }
                windowBase.SetValue(View.IsGeneratedProperty, true);
            }
            else
            {
                if (windowBase == null && !(model is IDisableWindowStyling))
                {
                    windowBase = new WindowBase
                    {
                        ResizeMode = resMode,
                        Content = view,
                    };

                    windowBase.SetValue(View.IsGeneratedProperty, true);
                    if (userControl != null)
                    {
                        windowBase.Width = userControl.ActualWidth;
                        windowBase.Height = userControl.ActualHeight;
                    }
                }

            }

            var window = base.EnsureWindow(model, windowBase ?? view, isDialog);

            //if (model is IHaveHelp)
               // Run.Safely(() => HelpProvider.SetContextId(window, ((IHaveHelp)model).HelpContextId));

            return window;
        }

        #endregion

        #region Implementation

        private readonly IDictionary<string, Window> _openedWindows = new Dictionary<string, Window>();

        private void ShowAndRegisterInOpenWindows(string windowId, object viewModel, Action<Window> configureWindow = null)
        {
            ShowWindow(
                viewModel,
                window =>
                {
                    _openedWindows.Add(windowId, window);
                    window.Closed += (s, e) =>
                    {
                        _openedWindows.Remove(windowId);
                        ((Window)s).Content = null;
                    };

                    if (configureWindow != null)
                        configureWindow(window);
                });
        }

        private void CloseWindow(string windowId)
        {
            Window window;
            if (_openedWindows.TryGetValue(windowId, out window))
            {
                // Paranoid removal in case that Closed event on window will not be hit for some reason.
                _openedWindows.Remove(windowId);
                Run.OnUIAndWait(window.Close);
            }
        }

        #endregion
    }
}
