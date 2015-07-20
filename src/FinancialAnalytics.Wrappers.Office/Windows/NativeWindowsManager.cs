using System;
using System.Runtime.InteropServices;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Windows
{
    public static class NativeWindowsManager
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("User32.dll", SetLastError = true)]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, ref IntPtr lParam);

        [DllImport("User32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool IsWindowEnabled(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hwnd, ref IntPtr lParam);

        public static IntPtr GetProcessId(IntPtr windowHandle)
        {
            uint result;
            GetWindowThreadProcessId(windowHandle, out result);
            return (IntPtr)result;
        }

        public static IntPtr GetChild(IntPtr mainWindowHandle, string childWindowClassName)
        {
            if (mainWindowHandle != IntPtr.Zero)
            {
                var childHandle = IntPtr.Zero;

                var enumChildren = new EnumWindowsProc((IntPtr currentChildHandle, ref IntPtr lParam) =>
                {
                    var buffer = new StringBuilder(128);
                    GetClassName(currentChildHandle, buffer, 128);
                    if (buffer.ToString() == childWindowClassName)
                    {
                        lParam = currentChildHandle;
                        return false;
                    }
                    return true;
                });

                EnumChildWindows(mainWindowHandle, enumChildren, ref childHandle);

                return childHandle;
            }
            return IntPtr.Zero;
        }
    }
}