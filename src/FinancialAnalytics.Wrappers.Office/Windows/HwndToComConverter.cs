using System;
using System.Runtime.InteropServices;

namespace FinancialAnalytics.Wrappers.Office.Windows
{
    public static class HwndToComConverter
    {
        [DllImport("Oleacc.dll")]
        internal static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, byte[] riid, out IDispatch ptr);

        internal const string IID_IDispatch = "00020400-0000-0000-C000-000000000046";
        internal const uint OBJID_NATIVEOM = 0xFFFFFFF0;

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid(IID_IDispatch)]
        internal interface IDispatch { }

        /// <summary>
        /// Returns window object from window handle, or null if no document is open.
        /// </summary>
        /// <remarks>
        /// Office application  Window class name  Result
        /// Word                _WwG               Window
        /// Excel               EXCEL7             Window
        /// PowerPoint          paneClassDC        DocumentWindow
        /// Command Bars        MsoCommandBar      CommandBar
        /// http://msdn.microsoft.com/en-us/library/dd317978(VS.85).aspx
        /// </remarks>
        public static object FindAcessibleObject(IntPtr mainWindowHandle, string documentWindowClassName)
        {
            var childHandle = NativeWindowsManager.GetChild(mainWindowHandle, documentWindowClassName);

            if (childHandle != IntPtr.Zero)
            {
                HwndToComConverter.IDispatch result;
                if (AccessibleObjectFromWindow(childHandle, OBJID_NATIVEOM, new Guid((string) IID_IDispatch).ToByteArray(), out result) >= 0)
                    return result;
            }

            return null;
        }
    }
}