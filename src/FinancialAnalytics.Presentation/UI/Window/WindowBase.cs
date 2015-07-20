using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Shapes;
using FinancialAnalytics.Presentation.Core;

namespace FinancialAnalytics.Presentation.UI.Window
{
    public class WindowBase : System.Windows.Window
    {
        private const int SYSTEM_COMM = 0x112;
        private HwndSource _wpfContent;

        public WindowBase()
        {
            var resourceProvider = new ResourceProvider();
            Resources.MergedDictionaries.Add(resourceProvider.GetResourceDictionary());

            Unloaded += WindowBase_Unloaded;
            SourceInitialized += InitWpfContent;
            PreviewMouseMove += ClearMousePointer;
            var windowBaseStyle = TryFindResource("WindowBase_Style") as Style;
            if (windowBaseStyle != null)
            {
                Style = windowBaseStyle;
            }
        }

        private void WindowBase_Unloaded(object sender, RoutedEventArgs e)
        {
            Resources.MergedDictionaries.Clear();
        }

        private void InitWpfContent(object sender, EventArgs e)
        {
            _wpfContent = PresentationSource.FromVisual((Visual)sender) as HwndSource;

            if (_wpfContent != null && _wpfContent.CompositionTarget != null)
            {
                _wpfContent.CompositionTarget.RenderMode = IsSoftwareModeOnly() ? RenderMode.SoftwareOnly : RenderMode.Default;
            }
        }

        private void ClearMousePointer(object s, MouseEventArgs eventArgs)
        {
            if (Mouse.LeftButton != MouseButtonState.Pressed)
            {
                Cursor = Cursors.Arrow;
            }
        }

        #region Click events
        protected void MinimizeClick(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        protected void RestoreClick(object sender, RoutedEventArgs e)
        {
            WindowState = (WindowState == WindowState.Normal) ? WindowState.Maximized : WindowState.Normal;
        }

        protected void CloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
        #endregion

        public override void OnApplyTemplate()
        {
            SetHeaderEvents();
            SetResizeEvents();

            base.OnApplyTemplate();
        }

        private void SetHeaderEvents()
        {
            var elementTitleBar = GetTemplateChild("PART_TitleBar") as UIElement;
            if (elementTitleBar != null)
            {
                elementTitleBar.MouseLeftButtonDown += delegate(object sender, MouseButtonEventArgs e)
                {
                    if (e.ClickCount == 2)
                    {
                        var templateChild = GetTemplateChild("PART_MaximizeButton") as Button;
                        if ((templateChild != null) && templateChild.IsVisible)
                        {
                            MaximizeOrRestoreWindow();
                        }
                    }
                    else
                    {
                        MoveWindow();
                    }
                };
            }

            //Add events to buttons
            var minimizeButton = GetTemplateChild("PART_MinimizeButton") as Button;
            if (minimizeButton != null)
                minimizeButton.Click += MinimizeClick;

            var maximizeButton = GetTemplateChild("PART_MaximizeButton") as Button;
            if (maximizeButton != null)
                maximizeButton.Click += RestoreClick;

            var closeButton = GetTemplateChild("PART_CloseButton") as Button;
            if (closeButton != null)
                closeButton.Click += CloseClick;
        }

        private void SetResizeEvents()
        {
            Rectangle rectLeft = base.GetTemplateChild("EXTERNAL_Border_Left") as Rectangle;
            if (rectLeft != null)
            {
                rectLeft.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectLeft.MouseMove += new MouseEventHandler(ShowProperPointer);
            }
            Rectangle rectRight = base.GetTemplateChild("EXTERNAL_Border_Right") as Rectangle;
            if (rectRight != null)
            {
                rectRight.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectRight.MouseMove += new MouseEventHandler(ShowProperPointer);
            }
            Rectangle rectTop = base.GetTemplateChild("EXTERNAL_Border_Top") as Rectangle;
            if (rectTop != null)
            {
                rectTop.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectTop.MouseMove += new MouseEventHandler(ShowProperPointer);
            }
            Rectangle rectBottom = base.GetTemplateChild("EXTERNAL_Border_Bottom") as Rectangle;
            if (rectBottom != null)
            {
                rectBottom.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectBottom.MouseMove += new MouseEventHandler(ShowProperPointer);
            }

            //angle


            Rectangle rectTopLeftTop = base.GetTemplateChild("EXTERNAL_Border_LeftTop") as Rectangle;
            if (rectTopLeftTop != null)
            {
                rectTopLeftTop.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectTopLeftTop.MouseMove += new MouseEventHandler(ShowProperPointer);
            }
            Rectangle rectTopLeft = base.GetTemplateChild("EXTERNAL_Border_TopLeft") as Rectangle;
            if (rectTopLeft != null)
            {
                rectTopLeft.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectTopLeft.MouseMove += new MouseEventHandler(ShowProperPointer);
            }

            //right top angle
            Rectangle rectRightTop = base.GetTemplateChild("EXTERNAL_Border_RightTop") as Rectangle;
            if (rectRightTop != null)
            {
                rectRightTop.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectRightTop.MouseMove += new MouseEventHandler(ShowProperPointer);
            }
            Rectangle rectTopRight = base.GetTemplateChild("EXTERNAL_Border_TopRight") as Rectangle;
            if (rectTopRight != null)
            {
                rectTopRight.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectTopRight.MouseMove += new MouseEventHandler(ShowProperPointer);
            }

            //right bottom angle
            Rectangle rectRightBottom = base.GetTemplateChild("EXTERNAL_Border_RightBottom") as Rectangle;
            if (rectRightBottom != null)
            {
                rectRightBottom.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectRightBottom.MouseMove += new MouseEventHandler(ShowProperPointer);
            }
            Rectangle rectBottomRight = base.GetTemplateChild("EXTERNAL_Border_BottomRight") as Rectangle;
            if (rectBottomRight != null)
            {
                rectBottomRight.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectBottomRight.MouseMove += new MouseEventHandler(ShowProperPointer);
            }

            //left bottom angle
            Rectangle rectBottomLeft = base.GetTemplateChild("EXTERNAL_Border_BottomLeft") as Rectangle;
            if (rectBottomLeft != null)
            {
                rectBottomLeft.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectBottomLeft.MouseMove += new MouseEventHandler(ShowProperPointer);
            }
            Rectangle rectLeftBottom = base.GetTemplateChild("EXTERNAL_Border_LeftBottom") as Rectangle;
            if (rectLeftBottom != null)
            {
                rectLeftBottom.PreviewMouseDown += new MouseButtonEventHandler(ChangeSize);
                rectLeftBottom.MouseMove += new MouseEventHandler(ShowProperPointer);
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private void ChangeWindowSize(SideOfModification direction)
        {
            SendMessage(_wpfContent.Handle, SYSTEM_COMM, (IntPtr)(61440 + direction), IntPtr.Zero);
        }

        private void ChangeSize(object sender, MouseButtonEventArgs eventArgs)
        {
            var clickedRegion = sender as Rectangle;
            if (clickedRegion == null)
                return;

            switch (clickedRegion.Name)
            {
                case "EXTERNAL_Border_Top":
                    Cursor = Cursors.SizeNS;
                    ChangeWindowSize(SideOfModification.Top);
                    break;
                case "EXTERNAL_Border_Bottom":
                    Cursor = Cursors.SizeNS;
                    ChangeWindowSize(SideOfModification.Bottom);
                    break;
                case "EXTERNAL_Border_Left":
                    Cursor = Cursors.SizeWE;
                    ChangeWindowSize(SideOfModification.Left);
                    break;
                case "EXTERNAL_Border_Right":
                    Cursor = Cursors.SizeWE;
                    ChangeWindowSize(SideOfModification.Right);
                    break;
                case "EXTERNAL_Border_TopLeft":
                    Cursor = Cursors.SizeNWSE;
                    ChangeWindowSize(SideOfModification.TopLeft);
                    break;
                case "EXTERNAL_Border_LeftTop":
                    Cursor = Cursors.SizeNWSE;
                    ChangeWindowSize(SideOfModification.TopLeft);
                    break;
                case "EXTERNAL_Border_TopRight":
                    Cursor = Cursors.SizeNESW;
                    ChangeWindowSize(SideOfModification.TopRight);
                    break;
                case "EXTERNAL_Border_RightTop":
                    Cursor = Cursors.SizeNESW;
                    ChangeWindowSize(SideOfModification.TopRight);
                    break;
                case "EXTERNAL_Border_LeftBottom":
                    Cursor = Cursors.SizeNESW;
                    ChangeWindowSize(SideOfModification.BottomLeft);
                    break;
                case "EXTERNAL_Border_BottomLeft":
                    Cursor = Cursors.SizeNESW;
                    ChangeWindowSize(SideOfModification.BottomLeft);
                    break;
                case "EXTERNAL_Border_BottomRight":
                    Cursor = Cursors.SizeNWSE;
                    ChangeWindowSize(SideOfModification.BottomRight);
                    break;
                case "EXTERNAL_Border_RightBottom":
                    Cursor = Cursors.SizeNWSE;
                    ChangeWindowSize(SideOfModification.BottomRight);
                    break;
            }
        }

        private void ShowProperPointer(object s, MouseEventArgs eventArgs)
        {
            var clickedRegion = s as Rectangle;
            if (clickedRegion == null)
                return;

            switch (clickedRegion.Name)
            {
                case "EXTERNAL_Border_Top":
                    Cursor = Cursors.SizeNS;
                    break;
                case "EXTERNAL_Border_Bottom":
                    Cursor = Cursors.SizeNS;
                    break;
                case "EXTERNAL_Border_Left":
                    Cursor = Cursors.SizeWE;
                    break;
                case "EXTERNAL_Border_Right":
                    Cursor = Cursors.SizeWE;
                    break;
                case "EXTERNAL_Border_TopLeft":
                    Cursor = Cursors.SizeNWSE;
                    break;
                case "EXTERNAL_Border_LeftTop":
                    Cursor = Cursors.SizeNWSE;
                    break;
                case "EXTERNAL_Border_TopRight":
                    Cursor = Cursors.SizeNESW;
                    break;
                case "EXTERNAL_Border_RightTop":
                    Cursor = Cursors.SizeNESW;
                    break;
                case "EXTERNAL_Border_LeftBottom":
                    Cursor = Cursors.SizeNESW;
                    break;
                case "EXTERNAL_Border_BottomLeft":
                    Cursor = Cursors.SizeNESW;
                    break;
                case "EXTERNAL_Border_RightBottom":
                    Cursor = Cursors.SizeNWSE;
                    break;
                case "EXTERNAL_Border_BottomRight":
                    Cursor = Cursors.SizeNWSE;
                    break;
            }
        }

        private static bool IsSoftwareModeOnly()
        {
            return false; //!PresentationSetup.UseHardwareAcceleration;
        }

        private void MaximizeOrRestoreWindow()
        {
            if (WindowState == WindowState.Normal)
            {
                WindowState = WindowState.Maximized;
            }
            else if (base.WindowState == WindowState.Maximized)
            {
                WindowState = WindowState.Normal;
            }
        }

        private void MoveWindow()
        {
            DragMove();
        }
    }

    public enum SideOfModification
    {
        Left = 0x01,
        Right = 0x02,
        Top = 0x03,
        TopLeft = 0x04,
        TopRight = 0x05,
        Bottom = 0x06,
        BottomLeft = 0x07,
        BottomRight = 0x08
    }
}
