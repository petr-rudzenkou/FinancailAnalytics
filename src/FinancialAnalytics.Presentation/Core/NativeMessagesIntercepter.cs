using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;

namespace FinancialAnalytics.Presentation.Core
{
    public class NativeMessagesInterceptor
    {
        private readonly Window _window;

        // Stored here to stop it from getting garbage collected
        private readonly Win32.MessageProcDelegate _messageProcDelegate;
        private readonly Win32.WndProcDelegate _wndProcDelegate;
        private readonly Win32.MouseProcDelegate _mouseProcDelegate;
        private IntPtr _hMessageProcHook;
        private IntPtr _hWindowProcHook;
        private IntPtr _hMouseProcHook;
        private IntPtr _windowHandle;


        private IntPtr _comboBoxPopupHandle;
        private readonly PopupData _popupData = new PopupData();
        private readonly Dictionary<IntPtr, SubclassedWindowData> _subclassedWindows = new Dictionary<IntPtr, SubclassedWindowData>();
        private bool _disposed;

        private class SubclassedWindowData
        {
            public IntPtr PrevHwndProc { get; set; }
            //we store this one just to prevent garbage collecting of the new Proc delegate
            public Win32.WndProc HookProcDelegate { get; set; }
        }

        //consolidates all data related to popup & its native window
        private class PopupData
        {
            public Popup Popup { get; set; }
            public IntPtr Handle { get; set; }
            //applicable only if popup has any ComboBox
            public bool InnerComboBoxIsOpen { get; set; }
            //if some popup is opened inside of our popup (is not applicable for combobox's popup)
            public IntPtr InnerPopupHandle { get; set; }
            //the element where popup naturally sets initial Mouse.Capture(), by default it is PopupRoot
            public FrameworkElement MouseCapturedElement { get; set; }
            public MouseButtonEventHandler PopupMouseUpHandler { get; set; }
            public MouseButtonEventHandler PreviewMouseDownOutsideCapturedElementHandler { get; set; }

            public void ReleaseReferencesAndUnsubscribe()
            {
                if (Popup != null)
                {
                    Popup.RemoveHandler(Mouse.PreviewMouseDownOutsideCapturedElementEvent,
                                        PreviewMouseDownOutsideCapturedElementHandler);
                    Popup.RemoveHandler(UIElement.MouseUpEvent, PopupMouseUpHandler);
                }
                Popup = null;
                Handle = IntPtr.Zero;
                MouseCapturedElement = null;
                InnerComboBoxIsOpen = false;
                InnerPopupHandle = IntPtr.Zero;
            }
        }

        public NativeMessagesInterceptor(Window window)
        {
            _window = window;

            _window.Closed += (sender, e) => Dispose();
            var hwndSource = (HwndSource)PresentationSource.FromVisual(window);
            _messageProcDelegate = MessageProcHook;
            _wndProcDelegate = WndProcHook;
            _mouseProcDelegate = MouseProcHook;
            _popupData.PopupMouseUpHandler = PopupMouseUp;
            _popupData.PreviewMouseDownOutsideCapturedElementHandler = PreviewMouseDownOutsideCapturedElement;

            if (hwndSource != null)
            {
                _windowHandle = hwndSource.Handle;
                CreateHooks();
            }
            else
                _window.SourceInitialized += OnWindowSourceInitialized;
        }

        /// <summary>
        /// Instructs interceptor to disregard processing of these key strokes, so they will go directly to Excel
        /// </summary>
        public IEnumerable<Key> SkipKeys { get; set; }

        private void OnWindowSourceInitialized(object sender, EventArgs e)
        {
            var source = (HwndSource)PresentationSource.FromVisual(_window);
            Debug.Assert(source != null);
            _windowHandle = source.Handle;

            CreateHooks();
        }

        private void CreateHooks()
        {
            uint threadId = Win32.GetWindowThreadProcessId(_windowHandle, IntPtr.Zero);
            _hMessageProcHook = Win32.SetWindowsHookEx(Win32.HookType.WH_GETMESSAGE, _messageProcDelegate, IntPtr.Zero, threadId);
            _hWindowProcHook = Win32.SetWindowsHookEx(Win32.HookType.WH_CALLWNDPROC, _wndProcDelegate, IntPtr.Zero, threadId);
            _hMouseProcHook = Win32.SetWindowsHookEx(Win32.HookType.WH_MOUSE, _mouseProcDelegate, IntPtr.Zero, threadId);
        }

        #region Mouse Hooking
        private int MouseProcHook(int nCode, IntPtr wParam, ref Win32.MouseHookStruct lParam)
        {
            switch (wParam.ToInt32())
            {
                case Win32.Messages.WM_LBUTTONDOWN:
                case Win32.Messages.WM_LBUTTONUP:
                case Win32.Messages.WM_LBUTTONDBLCLK:
                case Win32.Messages.WM_RBUTTONDOWN:
                case Win32.Messages.WM_RBUTTONUP:
                case Win32.Messages.WM_RBUTTONDBLCLK:
                case Win32.Messages.WM_MBUTTONDOWN:
                case Win32.Messages.WM_MBUTTONUP:
                case Win32.Messages.WM_MBUTTONDBLCLK:
                    if (lParam.hwnd != _windowHandle && lParam.hwnd != _popupData.Handle && lParam.hwnd != _comboBoxPopupHandle
                        && !Win32.IsChild(_windowHandle, lParam.hwnd)) //check for Win32.IsChild() is required when we have WindowsFormsHost somewhere in our window
                    //clicked outside of our window, we should deactivate it, so _window.IsActive will become false
                    {
                        if (_window.IsActive)
                        {
                            Win32.SendMessage(_windowHandle, Win32.Messages.WM_ACTIVATE, IntPtr.Zero, IntPtr.Zero);
                            //when clicked on other window, our textboxes shouldn't have focus anymore
                            Win32.SendMessage(_windowHandle, Win32.Messages.WM_KILLFOCUS, lParam.hwnd, IntPtr.Zero);

                            //BUG FIX IBTRDAO-3503. We must return focus to "XLMAIN", otherwise Excel refuses to handle mouse wheel etc.
                            var ownerWindowHandle = Win32.GetWindowLong(_windowHandle, Win32.WindowAttributes.GWL_HWNDPARENT);
                            if (ownerWindowHandle != IntPtr.Zero &&
                                (ownerWindowHandle == lParam.hwnd || Win32.IsChild(ownerWindowHandle, lParam.hwnd)))
                                Win32.SendMessage(ownerWindowHandle, Win32.Messages.WM_SETFOCUS, _windowHandle, IntPtr.Zero);
                        }
                    }
                    else if (lParam.hwnd == _windowHandle || Win32.IsChild(_windowHandle, lParam.hwnd))
                    {
                        //e.g. when clicked inside of TextBox in our window, the window should get activated to hook keyboard messages
                        if (!_window.IsActive)
                            Win32.SendMessage(_windowHandle, Win32.Messages.WM_ACTIVATE, new IntPtr(1), IntPtr.Zero);
                    }
                    break;
            }
            return Win32.CallNextHookEx(_hMessageProcHook, nCode, wParam, ref lParam);
        }
        #endregion

        #region Keyboard Hooking
        private int MessageProcHook(int nCode, IntPtr wParam, ref Win32.Message lParam)
        {
            if (nCode >= 0 && wParam.ToInt32() == 1)
                ProcessMessage(ref lParam);
            return Win32.CallNextHookEx(_hMessageProcHook, nCode, wParam, ref lParam);
        }

        private void ProcessMessage(ref Win32.Message message)
        {
            //check for Win32.IsChild() is required when we have WindowsFormsHost somewhere in our window
            if (_popupData.Handle.ToInt32() == 0 && _comboBoxPopupHandle.ToInt32() == 0 &&
                (!_window.IsActive || (message.hWnd != _windowHandle && !Win32.IsChild(_windowHandle, message.hWnd)))
                )
                return;

            switch (message.msg)
            {
                case Win32.Messages.WM_KEYDOWN: //0x100  
                case Win32.Messages.WM_KEYUP: //0x101 
                case Win32.Messages.WM_CHAR: //0x102  
                case Win32.Messages.WM_DEADCHAR: //0x103 
                case Win32.Messages.WM_SYSKEYDOWN: //0x104 
                case Win32.Messages.WM_SYSKEYUP: //0x105 
                case Win32.Messages.WM_SYSCHAR: //0x106  
                case Win32.Messages.WM_SYSDEADCHAR: //0x107 
                    Key key = KeyInterop.KeyFromVirtualKey(message.wparam.ToInt32());
                    if (SkipKeys == null || !SkipKeys.Contains(key))
                    {
                        var messageCopy = new Win32.Message { hWnd = message.hWnd, lparam = message.lparam, msg = message.msg, wparam = message.wparam };
                        //prevent further propagating of the source message, we don't want hosting environment
                        //to receive it (or it will start typing characters into Excel window)
                        message.msg = 0;
                        //since now there is noone else to translate key messages into character messages for us,
                        //we are doing it ourselves (WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP to WM_CHAR)
                        Win32.TranslateMessage(ref messageCopy);

                        //NOTE: we could've used ComponentDispatcher.RaiseThreadMessage(ref interopMsg); to notify WPF environment about our message
                        //NOTE: but some WinForms hosted contols and hosted WebBrowser still require explicit call of SendMessage()
                        //NOTE: Moreover, some key accelerators like Alt+F4 are getting messed up with ComponentDispatcher.RaiseThreadMessage, hence we'd better just resend this message
                        Win32.SendMessage(messageCopy.hWnd, messageCopy.msg, messageCopy.wparam, messageCopy.lparam);
                    }

                    break;
            }
        }
        #endregion

        #region Focus Hooking
        private int WndProcHook(int nCode, IntPtr wParam, ref Win32.CWPSTRUCT lParam)
        {
            if (nCode >= 0)
            {
                //ShowWindow(true)
                if (lParam.msg == Win32.Messages.WM_SHOWWINDOW && lParam.wparam.ToInt32() != 0)
                {
                    var hwndSource = HwndSource.FromHwnd(lParam.hWnd);
                    //check this is WPF popup
                    if (hwndSource != null && hwndSource.RootVisual.GetType().Name.Contains("PopupRoot"))
                    {
                        ControlType controlType;
                        Popup popup;
                        var parentHwndSource = GetParentHwndSource(hwndSource, out controlType, out popup);
                        if (parentHwndSource != null &&
                            (!popup.StaysOpen || controlType == ControlType.ComboBox ||
                            controlType == ControlType.ContextMenu || controlType == ControlType.MenuItem ||
                            controlType == ControlType.Other))
                        {
                            //we must handle only popups related to our Window or inner popups, not any others
                            if (parentHwndSource.Handle == _windowHandle)
                            {
                                switch (controlType)
                                {
                                    case ControlType.Other: //we don't need to handle Other popups, but need to handle ContextMenus inside such popups
                                        _popupData.ReleaseReferencesAndUnsubscribe();

                                        _popupData.Popup = popup;
                                        _popupData.Handle = lParam.hWnd;
                                        break;
                                    case ControlType.ContextMenu:
                                    case ControlType.CustomPopupContainingControl:
                                        RestoreWndProc(_popupData.Handle);
                                        _popupData.ReleaseReferencesAndUnsubscribe();

                                        _popupData.Popup = popup;
                                        _popupData.Handle = lParam.hWnd;
                                        SubclassWndProc(_popupData.Handle); //popup itself is getting bombed with WM_CAPTURECHANGED
                                        _popupData.MouseCapturedElement = Mouse.Captured as FrameworkElement ??
                                                                          hwndSource.RootVisual as FrameworkElement;
                                        popup.AddHandler(Mouse.PreviewMouseDownOutsideCapturedElementEvent,
                                                         _popupData.PreviewMouseDownOutsideCapturedElementHandler, true);
                                        popup.AddHandler(UIElement.MouseUpEvent, _popupData.PopupMouseUpHandler, true);
                                        break;
                                    case ControlType.MenuItem:
                                    case ControlType.ComboBox:
                                        _comboBoxPopupHandle = lParam.hWnd;
                                        SubclassWndProc(_windowHandle);
                                        //popup containing window is getting bombed with WM_CAPTURECHANGED
                                        break;
                                }
                            }
                            //comboboxes inside of popups are handled separately, so we are interested only in other types of inner popups
                            else if (parentHwndSource.Handle == _popupData.Handle && controlType != ControlType.ComboBox)
                            {
                                _popupData.InnerPopupHandle = lParam.hWnd;
                                SubclassWndProc(_popupData.InnerPopupHandle);
                                popup.AddHandler(Mouse.PreviewMouseDownOutsideCapturedElementEvent,
                                                 _popupData.PreviewMouseDownOutsideCapturedElementHandler, true);
                            }
                        }
                    }
                }
                //ShowWindow(false) - when ContextMenu closes we may come here
                else if (lParam.msg == Win32.Messages.WM_SHOWWINDOW &&
                    lParam.wparam.ToInt32() == 0 &&
                    _popupData.Handle == lParam.hWnd)
                {
                    RestoreWndProc(_popupData.Handle);
                }
                //DestroyWindow() - when ComboBox closes we'll come only here,
                //after ContextMenu is closed we'll also come here after ShowWindow(false), we are checking the hWnd
                else if (lParam.msg == Win32.Messages.WM_DESTROY &&
                    (_comboBoxPopupHandle == lParam.hWnd || _popupData.Handle == lParam.hWnd || _popupData.InnerPopupHandle == lParam.hWnd))
                {
                    if (_popupData.InnerPopupHandle == lParam.hWnd)
                    {
                        _subclassedWindows.Remove(_popupData.InnerPopupHandle);
                        _popupData.InnerPopupHandle = IntPtr.Zero;

                        if (ComponentDispatcher.IsThreadModal)
                        {
                            //Some modal window was opened
                            //We should mandatory Release Mouse (i.e. Mouse.Captured == null), otherwise some WPF windows will be unresponsible.
                            //Unfortunately, calling just Mouse.Capture(null) appears to be not enough, we should recapture the mouse to smth else before Releasing.
                            Mouse.Capture(_window);
                            Mouse.Capture(null);
                        }
                        else
                        {
                            var popup = _popupData.Popup;
                            if (popup != null && !popup.StaysOpen && _popupData.MouseCapturedElement != null)
                                Mouse.Capture(_popupData.MouseCapturedElement, CaptureMode.SubTree);
                        }
                    }
                    else
                    {
                        if (_comboBoxPopupHandle == lParam.hWnd)
                            RestoreWndProc(_windowHandle);

                        //We should mandatory Release Mouse (i.e. Mouse.Captured == null), otherwise some WPF windows will be unresponsible.
                        //Unfortunately, calling just Mouse.Capture(null) appears to be not enough, we should recapture the mouse to smth else before Releasing.
                        Mouse.Capture(_window);
                        Mouse.Capture(null);

                        _comboBoxPopupHandle = IntPtr.Zero;
                        _popupData.ReleaseReferencesAndUnsubscribe();
                        _subclassedWindows.Remove(lParam.hWnd);
                    }
                }
            }
            return Win32.CallNextHookEx(_hMessageProcHook, nCode, wParam, ref lParam);
        }

        #region Handling mouse clicks for opened popups

        private void PopupMouseUp(object sender, MouseButtonEventArgs e)
        {
            //if some modal window is opened from popup we shouldn't re-capture the mouse
            //see http://social.msdn.microsoft.com/Forums/en-US/wpf/thread/c95f1acb-5dee-4670-b779-b07b06afafff/
            if (ComponentDispatcher.IsThreadModal)
                return;

            if (_popupData.InnerComboBoxIsOpen || _popupData.InnerPopupHandle != IntPtr.Zero)
                return;

            var popup = _popupData.Popup;
            if (popup != null && !popup.StaysOpen && _popupData.MouseCapturedElement != null)
            {
                if (IsOpeningComboBox(e))
                    return;
                Mouse.Capture(_popupData.MouseCapturedElement, CaptureMode.SubTree);
            }
        }

        private bool IsOpeningComboBox(MouseButtonEventArgs e)
        {
            var toggleButton = ((FrameworkElement)e.OriginalSource).TemplatedParent as FrameworkElement;
            if (toggleButton != null)
            {
                var comboBox = toggleButton.TemplatedParent as ComboBox;
                if (comboBox != null)
                {
                    if (comboBox.IsDropDownOpen)
                    {
                        _popupData.InnerComboBoxIsOpen = true;
                    }
                    else
                    {
                        EventHandler comboBoxDropDownOpened = null;
                        comboBoxDropDownOpened = (s, args) =>
                        {
                            _popupData.InnerComboBoxIsOpen = true;
                            comboBox.DropDownOpened -= comboBoxDropDownOpened;
                        };
                        comboBox.DropDownOpened += comboBoxDropDownOpened;
                    }

                    EventHandler comboBoxDropDownClosed = null;
                    comboBoxDropDownClosed = (s, args) =>
                    {
                        _popupData.InnerComboBoxIsOpen = false;
                        if (_popupData.Popup != null && _popupData.Popup.IsOpen)
                            Mouse.Capture(_popupData.MouseCapturedElement, CaptureMode.SubTree);
                        else
                            Mouse.Capture(null);
                        comboBox.DropDownClosed -= comboBoxDropDownClosed;
                    };
                    comboBox.DropDownClosed += comboBoxDropDownClosed;
                    return true;
                }
            }
            return false;
        }

        private void PreviewMouseDownOutsideCapturedElement(object sender, MouseButtonEventArgs e)
        {
            //if some modal window is opened from popup we shouldn't handle clicks to close popup
            //see http://social.msdn.microsoft.com/Forums/en-US/wpf/thread/c95f1acb-5dee-4670-b779-b07b06afafff/
            if (ComponentDispatcher.IsThreadModal)
                return;

            var popup = _popupData.Popup;
            var capturedElement = _popupData.MouseCapturedElement;
            if (popup != null && !popup.StaysOpen && popup.IsOpen && capturedElement.InputHitTest(e.GetPosition(capturedElement)) == null)
                popup.IsOpen = false;
        }
        #endregion

        private enum ControlType
        {
            Other,
            ContextMenu,
            ComboBox,
            MenuItem,
            CustomPopupContainingControl
        }

        private HwndSource GetParentHwndSource(HwndSource popupBasedHwndSource, out ControlType controlType, out Popup popup)
        {
            controlType = ControlType.Other;

            var parentPopup = (Popup)((FrameworkElement)popupBasedHwndSource.RootVisual).Parent;
            popup = parentPopup;
            var parentHwndSource = (HwndSource)PresentationSource.FromVisual(parentPopup);
            if (parentHwndSource != null)
            {
                if (parentPopup.TemplatedParent != null || !parentPopup.StaysOpen)
                {
                    controlType = ControlType.CustomPopupContainingControl;

                    if (parentPopup.TemplatedParent is ComboBox)
                    {
                        controlType = ControlType.ComboBox;
                    }
                    if (parentPopup.TemplatedParent is MenuItem)
                    {
                        controlType = ControlType.MenuItem;
                    }
                }
            }
            else //NOTE: for ContextMenu & ToolTip - popup is not represented in the HwndSource
            {
                if (parentPopup.Child is ToolTip)
                    return null; //we don't want to handle ToolTip

                parentHwndSource = (HwndSource)PresentationSource.FromVisual(parentPopup.PlacementTarget);
                controlType = popup.Child is ContextMenu
                                ? ControlType.ContextMenu
                                : ControlType.CustomPopupContainingControl;
            }
            return parentHwndSource;
        }

        private IntPtr CustomWndProc(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam)
        {
            if (msg == Win32.Messages.WM_CAPTURECHANGED && lParam.ToInt32() == 0)
            {
                //Mouse.Capture(null, CaptureMode.None); //just to be a bit friendly with Excel's expectations
                if (!IsWpfCall())
                    return IntPtr.Zero;
            }
            if (!_subclassedWindows.ContainsKey(hWnd))
                return IntPtr.Zero;
            return Win32.CallWindowProc(_subclassedWindows[hWnd].PrevHwndProc, hWnd, msg, wParam, lParam);
        }

        /// <summary>
        /// Ad-hoc solution to distinguish between wpf and native calls. Performance critical.
        /// </summary>
        private static bool IsWpfCall()
        {
            const int MAX_STACK_EXAMINE_LENGTH = 15; //heuristic constant
            const string RELEASE_MOUSE_CAPTURE = "ReleaseMouseCapture";

            var stackFrames = new StackTrace().GetFrames();
            return stackFrames != null &&
                   stackFrames.Take(MAX_STACK_EXAMINE_LENGTH).Any(frame => frame.GetMethod().Name.Contains(RELEASE_MOUSE_CAPTURE));
        }

        private void SubclassWndProc(IntPtr hWnd)
        {
            if (_subclassedWindows.ContainsKey(hWnd))
                return;
            bool is64Bit = IntPtr.Size == 8;
            Win32.WndProc hookProcDelegate = CustomWndProc;
            IntPtr prevWndProc;
            if (is64Bit)
            {
                prevWndProc = Win32.SetWindowLongPtr(hWnd, Win32.WindowAttributes.GWL_WNDPROC, Marshal.GetFunctionPointerForDelegate(hookProcDelegate));
            }
            else
            {
                prevWndProc = (IntPtr)Win32.SetWindowLong(hWnd, Win32.WindowAttributes.GWL_WNDPROC, (int)Marshal.GetFunctionPointerForDelegate(hookProcDelegate));
            }
            _subclassedWindows.Add(hWnd, new SubclassedWindowData { PrevHwndProc = prevWndProc, HookProcDelegate = hookProcDelegate });
        }

        private void RestoreWndProc(IntPtr hWnd)
        {
            if (!_subclassedWindows.ContainsKey(hWnd))
                return;

            bool is64Bit = IntPtr.Size == 8;
            if (is64Bit)
            {
                Win32.SetWindowLongPtr(hWnd, Win32.WindowAttributes.GWL_WNDPROC, _subclassedWindows[hWnd].PrevHwndProc);
            }
            else
            {
                Win32.SetWindowLong(hWnd, Win32.WindowAttributes.GWL_WNDPROC, _subclassedWindows[hWnd].PrevHwndProc.ToInt32());
            }
            _subclassedWindows.Remove(hWnd);
        }
        #endregion

        #region IDisposable implementation

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~NativeMessagesInterceptor()
        {
            try
            {
                Dispose(false);
            }
            catch
            {
            }
        }

        private void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                _disposed = true;
                if (disposing)
                {
                    // Free managed resources here
                }
                // Free unmanaged resources here
                Win32.UnhookWindowsHookEx(_hMessageProcHook);
                Win32.UnhookWindowsHookEx(_hWindowProcHook);
                Win32.UnhookWindowsHookEx(_hMouseProcHook);
            }
        }

        #endregion

        #region Interop Stuff

        protected static class Win32
        {
            public class HookType
            {
                public const int WH_JOURNALRECORD = 0;
                public const int WH_JOURNALPLAYBACK = 1;
                public const int WH_KEYBOARD = 2;
                public const int WH_GETMESSAGE = 3;
                public const int WH_CALLWNDPROC = 4;
                public const int WH_CBT = 5;
                public const int WH_SYSMSGFILTER = 6;
                public const int WH_MOUSE = 7;
                public const int WH_HARDWARE = 8;
                public const int WH_DEBUG = 9;
                public const int WH_SHELL = 10;
                public const int WH_FOREGROUNDIDLE = 11;
                public const int WH_CALLWNDPROCRET = 12;
                public const int WH_KEYBOARD_LL = 13;
                public const int WH_MOUSE_LL = 14;
            }

            public class Messages
            {
                public const int WM_KEYDOWN = 0x100;
                public const int WM_KEYUP = 0x101;
                public const int WM_CHAR = 0x102;
                public const int WM_DEADCHAR = 0x103;
                public const int WM_SYSKEYDOWN = 0x104;
                public const int WM_SYSKEYUP = 0x105;
                public const int WM_SYSCHAR = 0x106;
                public const int WM_SYSDEADCHAR = 0x107;

                public const int WM_CAPTURECHANGED = 0x0215;
                public const int WM_SHOWWINDOW = 0x0018;
                public const int WM_DESTROY = 0x0002;
                public const int WM_ACTIVATE = 0x0006;

                public const int WM_LBUTTONDOWN = 0x201;
                public const int WM_LBUTTONUP = 0x202;
                public const int WM_LBUTTONDBLCLK = 0x203;
                public const int WM_RBUTTONDOWN = 0x204;
                public const int WM_RBUTTONUP = 0x205;
                public const int WM_RBUTTONDBLCLK = 0x206;
                public const int WM_MBUTTONDOWN = 0x207;
                public const int WM_MBUTTONUP = 0x208;
                public const int WM_MBUTTONDBLCLK = 0x209;

                public const int WM_KILLFOCUS = 0x0008;
                public const int WM_SETFOCUS = 0x0007;
            }

            public class WindowAttributes
            {
                public const int GWL_WNDPROC = -4;
                public const int GWL_HWNDPARENT = -8;
            }

            public struct Message
            {
                public IntPtr hWnd;
                public int msg;
                public IntPtr wparam;
                public IntPtr lparam;
            }

            public struct CWPSTRUCT
            {
                public IntPtr lparam;
                public IntPtr wparam;
                public int msg;
                public IntPtr hWnd;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct POINT
            {
                public int x;
                public int y;
            }

            //Declare the wrapper managed MouseHookStruct class.
            [StructLayout(LayoutKind.Sequential)]
            public struct MouseHookStruct
            {
                public POINT pt;
                public IntPtr hwnd;
                public int wHitTestCode;
                public int dwExtraInfo;
            }

            public delegate int MessageProcDelegate(int nCode, IntPtr wParam, ref Message m);
            public delegate int WndProcDelegate(int nCode, IntPtr wParam, ref CWPSTRUCT m);
            public delegate int MouseProcDelegate(int nCode, IntPtr wParam, ref MouseHookStruct m);
            public delegate IntPtr WndProc(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern IntPtr SetWindowsHookEx(int hookType, Delegate callback,
                                                         IntPtr hMod, uint dwThreadId);

            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern bool UnhookWindowsHookEx(IntPtr hhk);

            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern int CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, ref Message m);

            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern int CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, ref CWPSTRUCT m);

            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern int CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, ref MouseHookStruct m);

            [DllImport("coredll.dll", SetLastError = true)]
            public static extern IntPtr GetModuleHandle(string module);

            [DllImport("user32.dll", EntryPoint = "TranslateMessage")]
            public extern static bool TranslateMessage(ref Message m);

            [DllImport("user32.dll")]
            public extern static uint GetWindowThreadProcessId(IntPtr window, IntPtr module);

            [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
            public extern static int GetMessageTime();

            [DllImport("user32.dll")]
            public extern static IntPtr CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

            [DllImport("user32.dll")]
            public extern static int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

            [DllImport("user32.dll")]
            public extern static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            public extern static IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

            [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
            public static extern bool IsChild(IntPtr hWndParent, IntPtr hwnd);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr GetWindowLong(IntPtr hWnd, int nIndex);
        }

        #endregion
    }
}
