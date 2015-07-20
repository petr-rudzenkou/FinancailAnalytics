using Microsoft.Vbe.Interop;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class VbeWindowStateToWindowStateConverter
    {
        public static WindowState Convert(vbext_WindowState xlWindowState)
        {
            WindowState windowState;
            switch (xlWindowState)
            {
                case vbext_WindowState.vbext_ws_Maximize:
                    windowState = WindowState.Maximized;
                    break;
                case vbext_WindowState.vbext_ws_Minimize:
                    windowState = WindowState.Minimized;
                    break;
                default:
                    windowState = WindowState.Normal;
                    break;
            }
            return windowState;
        }

        public static vbext_WindowState ConvertBack(WindowState windowState)
        {
            vbext_WindowState xlWindowState;
            switch (windowState)
            {
                case WindowState.Maximized:
                    xlWindowState = vbext_WindowState.vbext_ws_Maximize;
                    break;
                case WindowState.Minimized:
                    xlWindowState = vbext_WindowState.vbext_ws_Minimize;
                    break;
                default:
                    xlWindowState = vbext_WindowState.vbext_ws_Normal;
                    break;
            }
            return xlWindowState;
        }
    }
}
