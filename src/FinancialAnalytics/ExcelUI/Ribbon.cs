using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using FinancialAnalytics.AuthenticationClient;
using FinancialAnalytics.DataFacades;
using FinancialAnalytics.ExcelUI.Ribbons;
using FinancialAnalytics.Presentation.Core;
using FinancialAnalytics.Presentation.Services;
using FinancialAnalytics.Views;
using Microsoft.Office.Core;

namespace FinancialAnalytics.ExcelUI
{
    public class Ribbon : RibbonBase, IRibbon
    {
        private IRibbonUI _ribbon;
        private readonly IViewsRenderer _viewsRenderer;
        private readonly IDataToolsRibbon _dataToolsRibbon;
        private readonly IAuthenticationClient _authenticationClient;

        public Ribbon(IViewsRenderer viewsRenderer, IDataToolsRibbon dataToolsRibbon, IAuthenticationClient authenticationClient)
        {
            _viewsRenderer = viewsRenderer;
            _dataToolsRibbon = dataToolsRibbon;
            _authenticationClient = authenticationClient;
            _authenticationClient.UserInfoUpdated += OnUserInfoUpdated;
            CreateRibbonElements();
        }

        public RefreshManager RefreshManager
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }

        public FinancialAnalytics.Views.ViewsRenderer ViewsRenderer
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }

        public bool ContainsElement(IRibbonControl control)
        {
            return RibbonElements.Any(x => x.Id == control.Id);
        }

        public void RibbonLoaded(IRibbonUI ribbon)
        {
            _dataToolsRibbon.RibbonLoaded(ribbon);
            _ribbon = ribbon;
        }

        public bool GetVisible(IRibbonControl control)
        {
            if (_dataToolsRibbon.ContainsElement(control))
            {
                return _dataToolsRibbon.GetVisible(control);
            }
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
            throw new NotImplementedException();
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            if (_dataToolsRibbon.ContainsElement(control))
            {
                return _dataToolsRibbon.GetImage(control);
            }
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
            if (_dataToolsRibbon.ContainsElement(control))
            {
                _dataToolsRibbon.OnToggleButtonAction(control, pressed);
                return;
            }
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
            if (_dataToolsRibbon.ContainsElement(control))
            {
                return _dataToolsRibbon.GetPressed(control);
            }
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
            if (_dataToolsRibbon.ContainsElement(control))
            {
                return _dataToolsRibbon.GetEnabled(control);
            }
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
            if (_dataToolsRibbon.ContainsElement(control))
            {
                return _dataToolsRibbon.GetContent(control);
            }
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
            if (_dataToolsRibbon.ContainsElement(control))
            {
                return _dataToolsRibbon.GetLabel(control);
            }
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
            if (_dataToolsRibbon.ContainsElement(control))
            {
                return _dataToolsRibbon.GetScreentip(control);
            }
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
            if (_dataToolsRibbon.ContainsElement(control))
            {
                return _dataToolsRibbon.GetSupertip(control);
            }
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
            if (_dataToolsRibbon.ContainsElement(control))
            {
                return _dataToolsRibbon.GetKeytip(control);
            }
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
                    //Main tab
                    new RibbonElement(RibbonIds.FINANCIAL_ANALYTICS_TAB,
                        () => Resources.RibbonResources.Ribbon_FinancialAnalytics_Tab),

                        //Get Data group
                        new RibbonElement(RibbonIds.FA_DATAGROUP,
                            () => Resources.RibbonResources.Ribbon_DataGroup_Label,
                            () => true,
                            () => true),
                            new RibbonElement(RibbonIds.FA_STOCK_SCREENER_BUTTON,
                                () => Resources.RibbonResources.Ribbon_StockScreener_Label,
                                () => _authenticationClient.IsOnline,
                                () => true,
                                () => _viewsRenderer.Show(ViewType.StockScreener),
                                (id) => Resources.RibbonResources.Deal_Screener_32x32px),
                                 new RibbonElement(RibbonIds.FA_QUOTES_BUTTON,
                                () => Resources.RibbonResources.Ribbon_Quotes_Button_Label,
                                () => _authenticationClient.IsOnline,
                                () => true,
                                () => _viewsRenderer.Show(ViewType.Quotes),
                                (id) => Resources.RibbonResources.Quotes_32x32px),
                            //new RibbonElement(RibbonIds.FA_LEAGUE_TABLE_BUTTON,
                            //    () => Resources.RibbonResources.Ribbon_LeagueTable_Label,
                            //    () => _authenticationClient.IsOnline,
                            //    () => true,
                            //    () => _viewsRenderer.Show(ViewType.LeagueTable),
                            //    (id) => Resources.RibbonResources.League_Table_32x32px),
                            new RibbonElement(RibbonIds.FA_HISTORICAL_DATA_BUTTON,
                                () => Resources.RibbonResources.Ribbon_HistoricalData_Button_Label,
                                () => _authenticationClient.IsOnline,
                                () => true,
                                () => _viewsRenderer.Show(ViewType.HistoricalData),
                                (id) => Resources.RibbonResources.history_data),
                                new RibbonElement(RibbonIds.FA_CHARTS_BUTTON,
                                () => Resources.RibbonResources.Ribbon_Charts_Button,
                                () =>  _authenticationClient.IsOnline,
                                () => true,
                                () => _viewsRenderer.Show(ViewType.Charts),
                                (id) => Resources.RibbonResources.charts_icon_32x32px),
                                //new RibbonElement(RibbonIds.FA_SEARCH_BUTTON,
                                //() => Resources.RibbonResources.Ribbon_Search_Button_Label,
                                //() =>  _authenticationClient.IsOnline,
                                //() => true,
                                //() => _viewsRenderer.Show(ViewType.Search),
                                //(id) => Resources.RibbonResources.Search_32x32px),
                                //Apps group
                        new RibbonElement(RibbonIds.FA_APPS_GROUP,
                            () => Resources.RibbonResources.Ribbon_AppsGroup_Label,
                            () => true,
                            () => true),
                            new RibbonElement(RibbonIds.FA_PORTFOLIO_BUTTON,
                            () => Resources.RibbonResources.Ribbon_Portfolio_Label,
                            () => _authenticationClient.IsOnline,
                            () => true,
                            () => _viewsRenderer.Show(ViewType.Portfolio),
                            (id) => Resources.RibbonResources.Portfolio_32x32px),
                            new RibbonElement(RibbonIds.FA_HOME_BUTTON,
                            () => Resources.RibbonResources.Ribbon_Home_Button,
                            () => true,
                            () => true,
                            () => Process.Start(Constants.HOME_PAGE_URL),
                            (id) => Resources.RibbonResources.Home_32x32px),
                         new RibbonElement(RibbonIds.FA_SETTINGS_GROUP,
                            () => Resources.RibbonResources.Ribbon_SettingsGroup_Label,
                            () => true,
                            () => true),
                            new RibbonElement(RibbonIds.FA_OPTIOINS_BUTTON,
                            () => Resources.RibbonResources.Ribbon_Options_Button_Label,
                            () => true,
                            () => true,
                            () => _viewsRenderer.Show(ViewType.Options),
                            (id) => Resources.RibbonResources.options_32x32px),

                            //new RibbonElement(RibbonIds.FA_STATUS_BUTTON,
                            //    () =>
                            //    {
                            //        if(_authenticationClient.IsOnline)
                            //           return Resources.RibbonResources.Ribbon_Logout_Button_Label;
                            //        return Resources.RibbonResources.Ribbon_Login_Button_Label;
                            //    },
                            //() => true,
                            //() => true,
                            //    () =>
                            //    {
                            //        if (_authenticationClient.IsOnline)
                            //        {
                            //            _authenticationClient.Logout();
                            //        }
                            //        else
                            //        {
                            //            _viewsRenderer.Show(ViewType.Login); 
                            //        }
                            //    },
                            //    (id) =>
                            //    {
                            //        if (_authenticationClient.IsOnline)
                            //            return Resources.RibbonResources.LogOut_16x16px;
                            //        return Resources.RibbonResources.login_16x16;
                            //    })

                            //new RibbonElement(RibbonIds.FA_LOGOUT_BUTTON,
                            //() => Resources.RibbonResources.Ribbon_Logout_Button_Label,
                            //() => true,
                            //() => true)
                });
            }
            catch (Exception ex)
            {
                //some logic of logging
            }
        }

        private void OnUserInfoUpdated(UserInfoArgs args)
        {
            _ribbon.Invalidate();
        }
    }
}
