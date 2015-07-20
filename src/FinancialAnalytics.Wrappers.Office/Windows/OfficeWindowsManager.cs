using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Windows
{
	public static class OfficeWindowsManager
    {
        private const string WordDocumentWindowClassName = "_WwG";
        private const string ExcelDocumentWindowClassName = "EXCEL7";
        private const string WordWindowClassName = "OpusApp";
        private const string ExcelWindowClassName = "XLMAIN";
        private const string VbaEditorWindowClassName = "wndclass_desked_gsk";
        private const string WordProcessName = "WINWORD";
        private const string ExcelProcessName = "EXCEL";
		private const string PowerPointProcessName = "POWERPNT";
		private static readonly List<string> PowerPointWindowClassNames = new List<string> { "PPTFrameClass", "PP12FrameClass", "PP11FrameClass", "PP10FrameClass", "PP9FrameClass" };
		private static readonly List<string> officeWindowsClassesNames;

		static OfficeWindowsManager()
		{
			officeWindowsClassesNames = new List<string> { WordWindowClassName, ExcelWindowClassName };
			officeWindowsClassesNames.AddRange(PowerPointWindowClassNames);
		}

		/// <summary>
        /// Returns Microsoft.Office.Interop.Word.Window
        /// </summary>
        public static object FindWordWindowObject(IntPtr mainWindowHandle)
        {
            return HwndToComConverter.FindAcessibleObject(mainWindowHandle, WordDocumentWindowClassName);
        }

		public static object FindPowerPointWindowObject(IntPtr mainWindowHandle)
		{
			foreach(var className in PowerPointWindowClassNames)
			{
				var obj = HwndToComConverter.FindAcessibleObject(mainWindowHandle, className);
				if (obj != null)
				{
					return obj;
				}
			}
			return null;
		}

        /// <summary>
        /// Returns Microsoft.Office.Interop.Excel.Window
        /// </summary>
        public static object FindExcelWindowObject(IntPtr mainWindowHandle)
        {
            return HwndToComConverter.FindAcessibleObject(mainWindowHandle, ExcelDocumentWindowClassName);
        }

		public static bool IsMainWindowHandle(IntPtr windowHandle)
		{
			StringBuilder buffer = new StringBuilder(128);
			NativeMethods.GetClassName(windowHandle, buffer, 128);

			return officeWindowsClassesNames.Contains(buffer.ToString());
		}

		public static IntPtr FindActiveMainWindowHandle()
		{

			IntPtr mainWindowHandle = NativeMethods.GetActiveWindow();

			return IsMainWindowHandle(mainWindowHandle) ? mainWindowHandle : IntPtr.Zero;
		}

		public static IntPtr FindWindow(Process process, bool? visible)
		{
			string processName = process.ProcessName.ToUpperInvariant();
			int processId = process.Id;
			switch (processName)
			{
				case ExcelProcessName:
					return FindWindow(ExcelWindowClassName, processId, visible);
				case WordProcessName:
					return FindWindow(WordWindowClassName, processId, visible);
				case PowerPointProcessName:
					return FindWindow(PowerPointWindowClassNames, processId, visible);
		}

			return IntPtr.Zero;
		}

		public static IntPtr FindWindow(Process process)
		{
			return FindWindow(process, null);
		}

		private static IntPtr FindWindow(string windowClassName, int processId, bool? visible)
		{
			IEnumerable<IntPtr> classNameMatches = TopWindowsFinder.Find(windowClassName);
			IntPtr result = classNameMatches.Where(hwnd =>
			{
				var windowProcessId = NativeWindowsManager.GetProcessId(hwnd).ToInt32();
				var match = processId == windowProcessId;
				bool visibleMatch = true;
				if (visible.HasValue)
				{
					visibleMatch = NativeMethods.IsWindowVisible(hwnd) == visible.Value;
				}
				return match && visibleMatch;
			}).FirstOrDefault();

			return result;
		}

		private static IntPtr FindWindow(IEnumerable<string> windowClassNames, int processId, bool? visible)
		{
			foreach (string windowClassName in windowClassNames)
			{
				IntPtr windowHandle = FindWindow(windowClassName, processId, visible);
				if (windowHandle != IntPtr.Zero)
				{
					return windowHandle;
				}
			}

			return IntPtr.Zero;
		}

    	public static IEnumerable<IntPtr> FindWordWindows()
        {
            return FindWindows(WordWindowClassName, WordProcessName);
        }

        public static IEnumerable<IntPtr> FindExcelWindows()
        {
            return FindWindows(ExcelWindowClassName, ExcelProcessName);
        }

		public static IEnumerable<IntPtr> FindPowerPointWindows()
		{
			List<IntPtr> windows = new List<IntPtr>();
			foreach (var className in PowerPointWindowClassNames)
			{
				windows.AddRange(FindWindows(className, PowerPointProcessName));
			}
			return windows;
		}

        public static IntPtr FindVbaEditorWindow()
        {
            var result = FindWindow(VbaEditorWindowClassName, Process.GetCurrentProcess().Id, null);
            return result;
        }

        /// <param name="window">Microsoft.Office.Interop.Word.Window</param>
        public static IntPtr FindWordWindowHandle(object window)
        {
            var result = FindWordWindows().FirstOrDefault(hwnd => FindWordWindowObject(hwnd) == window);
            return result;
        }

        private static IEnumerable<IntPtr> FindWindows(string windowClassName, string processName)
        {
            var classNameMatches = TopWindowsFinder.Find(windowClassName);
            var result = classNameMatches.Where(hwnd =>
            {
                var currentProcessId = NativeWindowsManager.GetProcessId(hwnd);
                var process = Process.GetProcessById((int)currentProcessId);
                var match = process.ProcessName.ToUpperInvariant() == processName.ToUpperInvariant();
                return match;
            });
            return result;
        }
    }
}