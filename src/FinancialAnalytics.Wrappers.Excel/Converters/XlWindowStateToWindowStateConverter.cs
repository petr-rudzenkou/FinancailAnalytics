using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlWindowStateToWindowStateConverter
    {
        public static WindowState Convert(XlWindowState xlWindowState)
        {
            WindowState windowState;
            switch (xlWindowState)
            {
                case XlWindowState.xlMaximized:
                    windowState = WindowState.Maximized;
                    break;
                case XlWindowState.xlMinimized:
                    windowState = WindowState.Minimized;
                    break;
                default:
                    windowState = WindowState.Normal;
                    break;
            }
            return windowState;
        }

        public static XlWindowState ConvertBack(WindowState windowState)
        {
            XlWindowState xlWindowState;
            switch (windowState)
            {
                case WindowState.Maximized:
                    xlWindowState = XlWindowState.xlMaximized;
                    break;
                case WindowState.Minimized:
                    xlWindowState = XlWindowState.xlMinimized;
                    break;
                default:
                    xlWindowState = XlWindowState.xlNormal;
                    break;
            }
            return xlWindowState;
        }
    }
}
