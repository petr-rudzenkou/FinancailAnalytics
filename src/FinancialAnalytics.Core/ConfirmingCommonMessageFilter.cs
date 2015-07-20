using System.Windows;
using FinancialAnalytics.Core.Notification;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Core
{
    public class ConfirmingCommonMessageFilter : StaComCrossThreadInvoker
    {
        private readonly IMessageBoxService _messageBoxService;

        public ConfirmingCommonMessageFilter(IMessageBoxService messageBoxService)
        {
            _messageBoxService = messageBoxService;
        }

        public ConfirmingCommonMessageFilter(uint totalWaitTime, IMessageBoxService messageBoxService)
            : base(totalWaitTime)
        {
            _messageBoxService = messageBoxService;
        }

        protected override bool ShouldRetry(uint totalWaitTime)
        {
            return _messageBoxService.Show(Resources.NotificationResources.The_application_is_still_busy_message, null, MessageBoxButton.YesNo)
                == MessageBoxResult.Yes;
        }
    }
}
