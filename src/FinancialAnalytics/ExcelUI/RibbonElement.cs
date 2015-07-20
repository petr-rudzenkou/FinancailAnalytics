using System;
using System.Collections.Generic;
using System.Drawing;

namespace FinancialAnalytics.ExcelUI
{
    public class RibbonElement : IRibbonElement
    {
        private readonly Func<bool> _getIsVisible;
        private readonly Func<string> _getLabel;
        private readonly Func<bool> _getIsEnabled;
        private readonly Func<string> _getScreenTip;
        private readonly Func<string> _getScreenSuperTip;
        private readonly Func<string> _getKeyTip;

        private readonly Action _action;
        private readonly Func<bool> _getIsPressed;
        private readonly Func<string> _getContent;

        private readonly Func<string, Bitmap> _getImage;
        private readonly Func<string, Bitmap> _getImagesMask;

        public RibbonElement(string id,
            Func<string> getLabel = null,
            Func<bool> getIsEnabled = null,
            Func<bool> getIsVisible = null,
            Action action = null,
            Func<string, Bitmap> getImage = null,
            Func<string, Bitmap> getImagesMask = null,
            Func<string> getScreenTip = null,
            Func<string> getScreenSuperTip = null,
            Func<string> getKeyTip = null,
            Func<bool> getIsPressed = null,
            Func<string> getContent = null)
        {
            Id = id;
            _getLabel = getLabel ?? (() => null);
            _getIsEnabled = getIsEnabled ?? (() => true);
            _getIsVisible = getIsVisible ?? (() => true);
            _getScreenTip = getScreenTip ?? (() => null);
            _getScreenSuperTip = getScreenSuperTip ?? (() => null);
            _getKeyTip = getKeyTip ?? (() => null);
            _action = action;
            _getIsPressed = getIsPressed ?? (() => false);
            _getContent = getContent ?? (() => null);
            _getImage = getImage;
            _getImagesMask = getImagesMask;
        }

        public string Id { get; private set; }

        public string Label
        {
            get { return _getLabel(); }
        }

        public bool IsEnabled
        {
            get { return _getIsEnabled(); }
        }

        public bool IsVisible
        {
            get { return _getIsVisible(); }
        }

        public string ScreenTip
        {
            get { return _getScreenTip(); }
        }

        public string ScreenSuperTip
        {
            get { return _getScreenSuperTip(); }
        }

        public string KeyTip
        {
            get { return _getKeyTip(); }
        }

        public Action Action
        {
            get { return _action; }
        }

        public bool IsPressed
        {
            get { return _getIsPressed(); }
        }

        public string Content
        {
            get { return _getContent(); }
        }

        public Bitmap Image
        {
            get { return _getImage(Id); }
        }

        public Bitmap ImagesMask
        {
            get { return _getImagesMask(Id); }
        }
    }
}
