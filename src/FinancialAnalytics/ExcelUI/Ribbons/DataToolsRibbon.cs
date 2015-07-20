using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using FinancialAnalytics.AuthenticationClient;
using FinancialAnalytics.Core.Formulas;
using FinancialAnalytics.Views;
using Microsoft.Office.Core;

namespace FinancialAnalytics.ExcelUI.Ribbons
{
    public class DataToolsRibbon : RibbonBase, IDataToolsRibbon
    {
        private readonly IViewsRenderer _viewsRenderer;
        private readonly IRefreshManager _refreshManager;
        private readonly IAuthenticationClient _authenticationClient;
        private readonly IRefreshFormulasTimer _refreshFormulasTimer;
        private readonly IDailyRefreshTimer _dailyRefreshTimer;

        private string _refreshMode = RibbonIds.FA_REFRESH_ALL_WORKBOOKS;
        private string _refreshButtonLabel = Resources.RibbonResources.Ribbon_Refresh_AllWorkbooks;
        private bool _autoUpdate = true;

        private IRibbonUI _ribbon;

        public DataToolsRibbon(IViewsRenderer viewsRenderer, IRefreshManager refreshManager, IRefreshFormulasTimer refreshFormulasTimer, IDailyRefreshTimer dailyRefreshTimer, IAuthenticationClient authenticationClient)
        {
            _viewsRenderer = viewsRenderer;
            _refreshManager = refreshManager;
            _refreshFormulasTimer = refreshFormulasTimer;
            _dailyRefreshTimer = dailyRefreshTimer;
            _refreshFormulasTimer.AutoUpdate += OnAutoUpdate;
            _dailyRefreshTimer.AutoUpdate += OnAutoUpdate;
            _authenticationClient = authenticationClient;
            CreateRibbonElements();
        }

        public bool ContainsElement(IRibbonControl control)
        {
            return RibbonElements.Any(x => x.Id == control.Id);
        }

        public void RibbonLoaded(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        public bool GetVisible(IRibbonControl control)
        {
            bool visible = false;
            var ribbonElement = FindRibbonElement(control);
            if (ribbonElement != null)
            {
                visible = ribbonElement.IsVisible;
            }
            return visible;
        }

        public bool GetMsoVisible(IRibbonControl control)
        {
            return true;
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            Bitmap image = null;
            var ribbonElement = FindRibbonElement(control);
            if (ribbonElement != null)
            {
                image = ribbonElement.Image;
            }
            return image;
        }

        public void OnToggleButtonAction(IRibbonControl control, bool pressed)
        {
            var ribbonElement = FindRibbonElement(control);
            if (ribbonElement != null)
            {
                ribbonElement.Action();
            }
        }

        public void OnButtonAction(IRibbonControl control)
        {
            OnToggleButtonAction(control, false);
        }

        public bool GetPressed(IRibbonControl control)
        {
            bool pressed = false;
            var ribbonElement = FindRibbonElement(control);
            if (ribbonElement != null)
            {
                pressed = ribbonElement.IsPressed;
            }
            return pressed;
        }

        public bool GetEnabled(IRibbonControl control)
        {
            bool isEnabled = false;
            var ribbonElement = FindRibbonElement(control);
            if (ribbonElement != null)
            {
                isEnabled = ribbonElement.IsEnabled;
            }
            return isEnabled;
        }

        public string GetContent(IRibbonControl control)
        {
            var content = string.Empty;
            var ribbonElement = FindRibbonElement(control);
            if (ribbonElement != null)
            {
                content = ribbonElement.Content;
            }
            return content;
        }

        public string GetLabel(IRibbonControl control)
        {
            var label = string.Empty;
            var ribbonElement = FindRibbonElement(control);
            if (ribbonElement != null)
            {
                label = ribbonElement.Label;
            }
            return label;
        }

        public string GetScreentip(IRibbonControl control)
        {
            var screenTip = string.Empty;
            var ribbonElement = FindRibbonElement(control);
            if (ribbonElement != null)
            {
                screenTip = ribbonElement.ScreenTip;
            }
            return screenTip;
        }

        public string GetSupertip(IRibbonControl control)
        {
            string superTip = string.Empty;
            var ribbonElement = FindRibbonElement(control);
            if (ribbonElement != null)
            {
                superTip = ribbonElement.ScreenSuperTip;
            }
            return superTip;
        }

        public Bitmap LoadImage(string itemId)
        {
            throw new NotImplementedException();
        }

        public int GetItemCount(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetItemId(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetItemLabel(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetSelectedItemId(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public int GetSelectedItemIndex(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public void OnChangeAction(Microsoft.Office.Core.IRibbonControl control, string selectedItemId, int selectedItemIndex)
        {
            throw new NotImplementedException();
        }

        public void OnChange(Microsoft.Office.Core.IRibbonControl control, string text)
        {
            throw new NotImplementedException();
        }

        public string GetText(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetKeytip(IRibbonControl control)
        {
            string keyTip = string.Empty;
            var ribbonElement = FindRibbonElement(control);
            if (ribbonElement != null)
            {
                keyTip = ribbonElement.KeyTip;
            }
            return keyTip;
        }

        private void CreateRibbonElements()
        {
            try
            {
                RibbonElements.AddRange(new List<IRibbonElement>
                {
                        //Data tools group
                        new RibbonElement(RibbonIds.FA_DATATOOLS_GROUP,
                            () => Resources.RibbonResources.Ribbon_DataToolsGroup,
                            () => true,
                            () => true),

                            //Refresh ribbon button
                            new RibbonElement(RibbonIds.FA_REFRESH_HOST,
                                () => string.Empty,
                                () => _authenticationClient.IsOnline,
                                () => true),
                            new RibbonElement(RibbonIds.FA_REFRESH_BUTTON,
                                () => _refreshButtonLabel,
                                () => true,
                                () => true,
                                () => _refreshManager.Refresh(_refreshMode),
                                (id) => Resources.RibbonResources.Refresh_Button_32x32px),
                                new RibbonElement(RibbonIds.FA_REFRESH_ALL_WORKBOOKS,
                                    () => Resources.RibbonResources.Ribbon_Refresh_AllWorkbooks,
                                    () => true,
                                    () => true,
                                    () => Refresh(RibbonIds.FA_REFRESH_ALL_WORKBOOKS),
                                    (id) =>
                                        IsRefreshModeChecked(id) ?
                                        Resources.RibbonResources.Refresh_Mode_Checked :
                                        null),
                                new RibbonElement(RibbonIds.FA_REFRESH_ACTIVE_WORKBOOK,
                                    () => Resources.RibbonResources.Ribbon_Refresh_ActiveWorkbook,
                                    () => true,
                                    () => true,
                                    () => Refresh(RibbonIds.FA_REFRESH_ACTIVE_WORKBOOK),
                                    (id) =>
                                        IsRefreshModeChecked(id) ?
                                        Resources.RibbonResources.Refresh_Mode_Checked :
                                        null),
                                new RibbonElement(RibbonIds.FA_REFRESH_ACTIVE_WORSHEET,
                                    () => Resources.RibbonResources.Ribbon_Refresh_ActiveWorksheet,
                                    () => true,
                                    () => true,
                                    () => Refresh(RibbonIds.FA_REFRESH_ACTIVE_WORSHEET),
                                    (id) =>
                                        IsRefreshModeChecked(id) ?
                                        Resources.RibbonResources.Refresh_Mode_Checked :
                                        null),
                                new RibbonElement(RibbonIds.FA_REFRESH_ACTIVE_CELL,
                                    () => Resources.RibbonResources.Ribbon_Refresh_ActiveCell,
                                    () => true,
                                    () => true,
                                    () => Refresh(RibbonIds.FA_REFRESH_ACTIVE_CELL),
                                    (id) =>
                                        IsRefreshModeChecked(id) ?
                                        Resources.RibbonResources.Refresh_Mode_Checked :
                                        null),
                                //new RibbonElement(RibbonIds.FA_UTILITIES_MENU,
                                //() => Resources.RibbonResources.Ribbon_Utilities_Menu_Label,
                                //() => true,
                                //() => true,
                                //null,
                                //(id) => Resources.RibbonResources.utilities_16x16px),
                                //new RibbonElement(RibbonIds.FA_UTILITIES_SHORTCUTS,
                                //() => Resources.RibbonResources.Ribbon_Utilities_Shortcuts,
                                //() => true,
                                //() => true,
                                //null,
                                //(id) => Resources.RibbonResources.shortcuts_16x16px),
                                 new RibbonElement(RibbonIds.FA_FA_XCHANGERATES_BUTTON,
                                () => Resources.RibbonResources.Ribbon_XchangeRates_Button_Label,
                                () => _authenticationClient.IsOnline,
                                () => true,
                                () => _viewsRenderer.Show(ViewType.XChangeRates),
                                (id) => Resources.RibbonResources.XchangeRates_16x16px),
                                new RibbonElement(RibbonIds.FA_UPDATE_STATUS_BUTTON,
                                () => _autoUpdate ? Resources.RibbonResources.Ribbon_Update_Running_Label : Resources.RibbonResources.Ribbon_Update_Paused_Label,
                                () => true,
                                () => true,
                                () => ToggleUpdate(),
                                (id) => _autoUpdate ? Resources.RibbonResources.pause_16px : Resources.RibbonResources.play_16px)

                });
            }
            catch (Exception ex)
            { }
        }

        private bool IsRefreshModeChecked(string elementId)
        {
            return _refreshMode.Equals(elementId);
        }

        private void Refresh(string refreshMode)
        {
            _refreshMode = refreshMode;

            switch (_refreshMode)
            {
                case RibbonIds.FA_REFRESH_BUTTON:
                {
                    _refreshButtonLabel = Resources.RibbonResources.Ribbon_Refresh_Button;
                    break;
                }
                case RibbonIds.FA_REFRESH_ALL_WORKBOOKS:
                {
                    _refreshButtonLabel = Resources.RibbonResources.Ribbon_Refresh_AllWorkbooks;
                    break;
                }
                case RibbonIds.FA_REFRESH_ACTIVE_WORKBOOK:
                {
                    _refreshButtonLabel = Resources.RibbonResources.Ribbon_Refresh_ActiveWorkbook;
                    break;
                }
                case RibbonIds.FA_REFRESH_ACTIVE_WORSHEET:
                {
                    _refreshButtonLabel = Resources.RibbonResources.Ribbon_Refresh_ActiveWorksheet;
                    break;
                }
                case RibbonIds.FA_REFRESH_ACTIVE_CELL:
                {
                    _refreshButtonLabel = Resources.RibbonResources.Ribbon_Refresh_ActiveCell;
                    break;
                }
                default:
                {
                    _refreshButtonLabel = Resources.RibbonResources.Ribbon_Refresh_Button;
                    break;
                }
            }

            _refreshManager.Refresh(_refreshMode);
            InvalidateRefreshMenu();
        }

        private void InvalidateRefreshMenu()
        {
            _ribbon.InvalidateControl(RibbonIds.FA_REFRESH_BUTTON);
            _ribbon.InvalidateControl(RibbonIds.FA_REFRESH_ALL_WORKBOOKS);
            _ribbon.InvalidateControl(RibbonIds.FA_REFRESH_ACTIVE_WORKBOOK);
            _ribbon.InvalidateControl(RibbonIds.FA_REFRESH_ACTIVE_WORSHEET);
            _ribbon.InvalidateControl(RibbonIds.FA_REFRESH_ACTIVE_CELL);
        }

        private void ToggleUpdate()
        {
            _autoUpdate = !_autoUpdate;
            if (_autoUpdate)
            {
                _refreshFormulasTimer.Start();
                _dailyRefreshTimer.Start();
            }
            else
            {
                _refreshFormulasTimer.Stop();
                _dailyRefreshTimer.Stop();
            }
            _ribbon.InvalidateControl(RibbonIds.FA_UPDATE_STATUS_BUTTON);
        }

        private void OnAutoUpdate(object sender, EventArgs e)
        {
            _refreshManager.Refresh(_refreshMode);
        }
    }
}
