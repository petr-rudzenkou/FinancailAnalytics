using System;
using System.Runtime.InteropServices;
using System.Text;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Office
{
    public class NativeMethods
    {
        private const int Minimize = 6;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool MoveWindow(IntPtr windowHandle, int x, int y, int width, int height, bool repaint);

		[DllImport("user32.dll")]
		private static extern Int32 GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("Ole32.dll")]
        public static extern int
            CoRegisterMessageFilter(IMessageFilter newFilter, out 
			IMessageFilter oldFilter);



        [DllImport("mpr.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int WNetGetConnection(
            [MarshalAs(UnmanagedType.LPTStr)] string localName,
            [MarshalAs(UnmanagedType.LPTStr)] StringBuilder remoteName,
            ref int length);


        public static bool MinimizeWindow(IntPtr hWnd)
        {
            return ShowWindow(hWnd, Minimize);
        }

        public static string GetClassName(IntPtr hWnd)
        {
            StringBuilder className = new StringBuilder(100);
            GetClassName(hWnd, className, className.Capacity);
            return className.ToString();
        }

		public static uint GetWindowThreadProcessId(IntPtr hWnd)
		{
			try
			{
				uint processId;
				GetWindowThreadProcessId(hWnd, out processId);
				return processId;
			}
			catch { return 0; }
		}
    }
}
