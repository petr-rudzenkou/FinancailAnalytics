using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using DryTools;
using DryTools.Execution;
using DryTools.Primitives;
using FinancialAnalytics.Core;
using FinancialAnalytics.Core.Composition.Unity;

namespace FinancialAnalytics.Presentation.Core
{
    public static class ConfigureWindow
    {
        private readonly static IApplicationProvider _applicationProvider;
        static ConfigureWindow()
        {
            _applicationProvider = Locator.Current.GetInstance<IApplicationProvider>();
        }
        public const double DEFAULT_MIN_WIDTH = 870;

        public const double DEFAULT_MIN_HEIGHT = 700;

        public static WindowStartupLocation DefaultWindowStartupLocation
        {
            get { return _defaultWindowStartupLocation; }
        }

        public static IntPtr TryGetWindowHandleOfView(object view)
        {
            if (view == null)
                return IntPtr.Zero;

            var viewElement = Ensure.OfType<UIElement>(() => view);

            var window = Window.GetWindow(viewElement);
            if (window == null)
                return IntPtr.Zero;

            var windowHandle = new WindowInteropHelper(window).Handle;
            if (windowHandle == IntPtr.Zero)
                return IntPtr.Zero;

            if (!IsWindowHandleValid(windowHandle))
                return IntPtr.Zero;

            return windowHandle;
        }

        public static Await<None> ShowModeless(this Window window, Action<NativeMessagesInterceptor> configureInterceptor = null)
        {
            window.WindowStartupLocation = DefaultWindowStartupLocation;
            window.MakeProcessOwned().Show();

            FixKeyboardAndMouseOverExcel(window, configureInterceptor);

            return Await.When(
                e => e.HasResult,
                h => window.Closed += h,
                h => window.Closed -= h)
                .Take();
        }

        public static Window DockTo(this Window window, FrameworkElement element)
        {
            try
            {
                bool isOnLeftSide;
                bool isOnTopSide;

                var childPointCenter = element.PointToScreen(new Point(element.ActualWidth / 2D, element.ActualHeight / 2D));
                var screenDimensions = SystemParameters.WorkArea;
                var source = PresentationSource.FromVisual(element);
                var targetChildPoints = source.CompositionTarget.TransformFromDevice.Transform(childPointCenter);

                Point point;
                if (targetChildPoints.X > ((screenDimensions.Left + screenDimensions.Right) / 2D))
                {
                    isOnLeftSide = true;
                    point = element.PointToScreen(new Point(0, 0));
                }
                else
                {
                    isOnLeftSide = false;
                    point = element.PointToScreen(new Point(element.ActualWidth, 0));
                }

                if (targetChildPoints.Y > ((screenDimensions.Top + screenDimensions.Bottom) / 2D))
                {
                    isOnTopSide = true;
                    point.Y = element.PointToScreen(new Point(0, 0)).Y;
                }
                else
                {
                    point.Y = element.PointToScreen(new Point(0, element.ActualHeight)).Y;
                    isOnTopSide = false;
                }

                var targetPoints = source.CompositionTarget.TransformFromDevice.Transform(point);

                window.ShowInTaskbar = false;
                window.WindowStartupLocation = WindowStartupLocation.Manual;

                window.Left = SystemParameters.VirtualScreenWidth;
                window.Top = SystemParameters.WindowCaptionHeight;
                var left = targetPoints.X;
                var top = targetPoints.Y;

                var dialogWidth = window.ActualWidth == 0d ? window.Width : window.ActualWidth;
                var dialogHeight = window.ActualHeight == 0d ? window.Height : window.ActualHeight;

                if (isOnLeftSide)
                {
                    left = left - dialogWidth;
                    if (screenDimensions.Left > left)
                    {
                        //Move to right side of the parent
                        left = left + dialogWidth + element.ActualWidth;
                        if ((left + dialogWidth) > screenDimensions.Right)
                        {
                            left = screenDimensions.Left;
                        }
                    }
                }
                else
                {
                    if (screenDimensions.Right < (left + dialogWidth))
                    {
                        //Move to left side of the parent
                        left = left - dialogWidth - element.ActualWidth;
                        if (left < screenDimensions.Left)
                        {
                            left = screenDimensions.Right - dialogWidth;
                        }
                    }
                }

                if (isOnTopSide)
                {
                    double actualHeight = dialogHeight - element.ActualHeight;
                    top = top - actualHeight;
                    if (screenDimensions.Top > top)
                    {
                        //Move to top side of the parent
                        top = top + actualHeight + element.ActualHeight;
                        if ((top + actualHeight) > screenDimensions.Bottom)
                        {
                            top = screenDimensions.Top;
                        }
                    }
                }
                else
                {
                    top = top - element.ActualHeight;
                    double actualHeight = dialogHeight + element.ActualHeight;
                    if (screenDimensions.Bottom < (window.Top + actualHeight))
                    {
                        //Move to bottom side of the parent
                        top = top - actualHeight - element.ActualHeight;
                        if (top < screenDimensions.Top)
                        {
                            top = screenDimensions.Bottom - actualHeight;
                        }
                    }
                }

                window.Left = left;
                window.Top = top;
                window.Visibility = Visibility.Visible;
            }
            catch (Exception e)
            {
                //Logging is not supported - throw exceptions as this is a stand alone library of controls
                throw new Exception("Error while dialog opening", e);
                //return new ColorDialogResult(ColorDialogResultType.Cancel, Colors.White);
            }

            return window;
        }

        public static Window MakeProcessOwned(this Window window)
        {
            window.ShowInTaskbar = false;

            //NOTE: Process.GetCurrentProcess().MainWindowHandle can return invalid handle for Excel.
            //NOTE: pulling out Excel's main window by WinAPI's ClassName, see http://msdn.microsoft.com/en-us/library/bb687833.aspx
            //var currentThreadId = GetCurrentThreadId();
            //var ownerHandle = FindWindowInThread(currentThreadId, className => "XLMAIN".Equals(className, StringComparison.OrdinalIgnoreCase)); doesn't work for 64 bit office

            var ownerHandle = _applicationProvider.GetIfReady(a => new IntPtr(a.WindowHandle), IntPtr.Zero);
            if (ownerHandle != IntPtr.Zero &&
                // Sometimes 'Invalid window handle' is raised.
                // This code is just checking this situation but does not fix the reason
                IsWindowHandleValid(ownerHandle))
            {
                SetForegroundWindow(ownerHandle);
                var windowInteropHelper = new WindowInteropHelper(window);
                windowInteropHelper.Owner = ownerHandle;
                window.Closing += (sender, args) => SetForegroundWindow(ownerHandle); //required when there are several opened windows & you're closing them
            }
            else
            {
                window.Owner = null;
            }

            return window;
        }
        public static IntPtr GetHandleId(this Window window)
        {
            return _applicationProvider.GetIfReady(a => new IntPtr(a.WindowHandle), IntPtr.Zero);
        }
        public static Window DisableParent(this Window window, object parent)
        {
            var parentHandle = parent is IntPtr ? (IntPtr)parent : TryGetWindowHandleOfView(parent);
            window.DisableParent(parentHandle);
            return window;
        }

        public static Window DisableParent(this Window window, IntPtr windowHandle)
        {
            if (windowHandle != IntPtr.Zero)
            {
                EnableWindow(windowHandle, false);
                window.Closed += (sender, args) => EnableWindow(windowHandle, true);
            }

            return window;
        }

        public static Window WithGrip(this Window window)
        {
            window.ResizeMode = ResizeMode.CanResizeWithGrip;
            window.MinHeight = DEFAULT_MIN_HEIGHT;
            window.MinWidth = DEFAULT_MIN_WIDTH;
            return window;
        }

        // NOTE: Garbage collection in the module can cause deadlock on the stage of module's unloading
        public static Window CleanMemoryAfterClosing(this Window window)
        {
            window.Unloaded += (s, e) => Run.Async(() =>
            {
                Thread.Sleep(1000);//At this moment there is no better way to perform operation after window's closing
                GC.Collect();
            });

            return window;
        }

        #region Implementation

        private const WindowStartupLocation _defaultWindowStartupLocation = WindowStartupLocation.CenterScreen;

        private static bool IsWindowHandleValid(IntPtr handle)
        {
            var info = new WINDOWINFO();
            info.cbSize = (uint)Marshal.SizeOf(info);
            return GetWindowInfo(handle, ref info);
        }

        private static void FixKeyboardAndMouseOverExcel(Window window, Action<NativeMessagesInterceptor> configureInterceptor = null)
        {
            var interceptor = new NativeMessagesInterceptor(window);
            window.Closed += (sender, args) => interceptor.Dispose();
            if (configureInterceptor != null)
                configureInterceptor(interceptor);
        }

        #endregion

        #region Native methods

        [DllImport("User32.dll")]
        public static extern int SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowInfo(IntPtr hwnd, ref WINDOWINFO pwi);

        [StructLayout(LayoutKind.Sequential)]
        struct WINDOWINFO
        {
            public uint cbSize;
            public RECT rcWindow;
            public RECT rcClient;
            public uint dwStyle;
            public uint dwExStyle;
            public uint dwWindowStatus;
            public uint cxWindowBorders;
            public uint cyWindowBorders;
            public ushort atomWindowType;
            public ushort wCreatorVersion;

            public WINDOWINFO(Boolean? filler)
                : this()   // Allows automatic initialization of "cbSize" with "new WINDOWINFO(null/true/false)".
            {
                cbSize = (UInt32)(Marshal.SizeOf(typeof(WINDOWINFO)));
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            private int _Left;
            private int _Top;
            private int _Right;
            private int _Bottom;
        }

        //Doesn't work for 64 bit office
        //private static IntPtr FindWindowInThread(int threadId, Func<string, bool> compareClassName)
        //{
        //    IntPtr windowHandle = IntPtr.Zero;
        //    EnumThreadWindows(threadId, (IntPtr hWnd, ref IntPtr lParam) =>
        //    {
        //        var className = new StringBuilder(200);
        //        GetClassName(hWnd, className, 200);
        //        if (compareClassName(className.ToString()))
        //        {
        //            windowHandle = hWnd;
        //            return false;
        //        }
        //        return true;
        //    }, IntPtr.Zero);

        //    return windowHandle;
        //}

        #endregion
    }
}
