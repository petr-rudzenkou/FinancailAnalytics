using System;
using System.Windows;
using FinancialAnalytics.Presentation.Core;
using FinancialAnalytics.Views.ViewSettings;

namespace FinancialAnalytics.Views
{
    //Consider using presentation serivce to render views
    public class ViewsRenderer : IViewsRenderer
    {
        private readonly IOfficeWindowManager _officeWindowManager;
        private readonly IWindowSettingsFactory _windowSettingsFactory;
        private readonly IViewModelFactory _viewModelFactory;

        public ViewsRenderer(IOfficeWindowManager officeWindowManager, IWindowSettingsFactory windowSettingsFactory, IViewModelFactory viewModelFactory)
        {
            _officeWindowManager = officeWindowManager;
            _windowSettingsFactory = windowSettingsFactory;
            _viewModelFactory = viewModelFactory;
        }
        public void Show(ViewType viewType)
        {
            var viewModel = _viewModelFactory.Create(viewType);
            var windowId = viewModel.GetType().FullName;
            if (viewType == ViewType.Login)
            {
                _officeWindowManager.ShowDialog(viewModel, w => ConfigureWindow(w, viewType));
                return;
            }
           ShowWindow(windowId, viewModel, w => ConfigureWindow(w, viewType));
        }

        private void ShowWindow(string windowId, object viewModel, Action<Window> configureWindow = null)
        {
            if (!_officeWindowManager.ShowWindowOnlyWhenPreviousIsClosed(windowId, viewModel, configureWindow))
            {
                _officeWindowManager.TryActivateWindow(windowId);
            }
        }

        private void ConfigureWindow(Window window, ViewType viewType)
        {
            var windowSettings = _windowSettingsFactory.GetWindowSettings(viewType);
            if (windowSettings != null)
            {
                if (windowSettings.Height != null)
                    window.Height = windowSettings.Height.Value;
                if (windowSettings.Width != null)
                    window.Width = windowSettings.Width.Value;
                if (windowSettings.ResizeMode != null)
                    window.ResizeMode = windowSettings.ResizeMode.Value;
                window.WindowStyle = WindowStyle.None;
            }
        }
    }
}
