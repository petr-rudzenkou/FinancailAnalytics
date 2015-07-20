using System.Windows;

namespace FinancialAnalytics.Core.Notification
{
    public interface IMessageBoxService
    {
        /// <summary>
        /// Show message box with specific content, default image, default set of buttons, without title and details 
        /// </summary>
        /// <param name="content">Main text for message box</param>
        /// <returns>Result of user's response to message box</returns>
        MessageBoxResult Show(string content);

        /// <summary>
        /// Show message box with specific content and details, default image, default set of buttons, without title
        /// </summary>
        /// <param name="content">Main text for message box</param>
        /// <param name="details">Additional details text</param>
        /// <returns>Result of user's response to message box</returns>
        MessageBoxResult Show(string content, string details);

        /// <summary>
        /// Show message box with specific content, details, image type and default set of buttons, without title
        /// </summary>
        /// <param name="content">Main text for message box</param>
        /// <param name="details">Additional details text</param>
        /// <param name="imageType">Type of image which should be displayed in the message box</param>
        /// <returns>Result of user's response to message box</returns>
        MessageBoxResult Show(string content, string details, MessageBoxImage imageType);

        /// <summary>
        /// Show message box with specific content, details, set of buttons and default image type, without title
        /// </summary>
        /// <param name="content">Main text for message box</param>
        /// <param name="details">Additional details text</param>
        /// <param name="buttonsSet">Set of buttons which should be available in the message box</param>
        /// <returns>Result of user's response to message box</returns>
        MessageBoxResult Show(string content, string details, MessageBoxButton buttonsSet);

        /// <summary>
        /// Show message box with specific content, details, image type, set of buttons and without title
        /// </summary>
        /// <param name="content">Main text for message box</param>
        /// <param name="details">Additional details text</param>
        /// <param name="imageType">Type of image which should be displayed in the message box</param>
        /// <param name="buttonsSet">Set of buttons which should be available in the message box</param>
        /// <returns>Result of user's response to message box</returns>
        MessageBoxResult Show(string content, string details, MessageBoxImage imageType, MessageBoxButton buttonsSet);

        /// <summary>
        /// Show message box with specific title, content and details, default image and default set of buttons
        /// </summary>
        /// <param name="title">Title for message box</param>
        /// <param name="content">Main text for message box</param>
        /// <param name="details">Additional details text</param>
        /// <returns>Result of user's response to message box</returns>
        MessageBoxResult Show(string title, string content, string details);

        /// <summary>
        /// Show message box with specific title, content, details, image type and default set of buttons
        /// </summary>
        /// <param name="title">Title for message box</param>
        /// <param name="content">Main text for message box</param>
        /// <param name="details">Additional details text</param>
        /// <param name="imageType">Type of image which should be displayed in the message box</param>
        /// <returns>Result of user's response to message box</returns>
        MessageBoxResult Show(string title, string content, string details, MessageBoxImage imageType);

        /// <summary>
        /// Show message box with specific title, content, details, set of buttons and default image type
        /// </summary>
        /// <param name="title">Title for message box</param>
        /// <param name="content">Main text for message box</param>
        /// <param name="details">Additional details text</param>
        /// <param name="buttonsSet">Set of buttons which should be available in the message box</param>
        /// <returns>Result of user's response to message box</returns>
        MessageBoxResult Show(string title, string content, string details, MessageBoxButton buttonsSet);

        /// <summary>
        /// Show message box with specific title, content, details, image type, set of buttons
        /// </summary>
        /// <param name="title">Title for message box</param>
        /// <param name="content">Main text for message box</param>
        /// <param name="details">Additional details text</param>
        /// <param name="imageType">Type of image which should be displayed in the message box</param>
        /// <param name="buttonsSet">Set of buttons which should be available in the message box</param>
        /// <returns>Result of user's response to message box</returns>
        MessageBoxResult Show(string title, string content, string details, MessageBoxImage imageType, MessageBoxButton buttonsSet);
    }
}
