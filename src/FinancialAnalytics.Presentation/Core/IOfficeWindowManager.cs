using System;
using System.Windows;

namespace FinancialAnalytics.Presentation.Core
{
    public interface IOfficeWindowManager
    {
        void ShowPopup(object rootModel);

        bool? ShowDialog(object rootModel, Action<Window> configureWindow = null);

        void ShowWindow(object viewModel, Action<Window> configureWindow = null, Action<NativeMessagesInterceptor> configureInterceptor = null);

        /// <returns>false - if window is already opened, true - if was opened in this method call</returns>
        bool ShowWindowOnlyWhenPreviousIsClosed(string windowId, object viewModel, Action<Window> configureWindow = null);

        void ShowWindowAndClosePrevious(string windowId, object viewModel, Action<Window> configureWindow = null);

        bool IsWindowActive(string windowId);

        void CloseAllWindows();

        bool TryRegisterWindow(string windowId, Window window);

        bool TryActivateWindow(string windowId);
    }
}
