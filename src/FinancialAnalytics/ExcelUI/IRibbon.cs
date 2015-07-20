using System.Drawing;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.ExcelUI
{
    public interface IRibbon
    {
        /// <summary>
        /// Callback invoked when custom UI is loaded.
        /// </summary>
        /// <param name="ribbon">
        /// IRibbon object is passed as a parameter. This object exposes Invalidate and InvalidateControl.
        /// </param>
        void RibbonLoaded(IRibbonUI ribbon);

        /// <summary>
        /// Callback that returns true if the control is visible. By default this method should return false
        /// </summary>
        /// <param name="control">
        /// The ribbon control.
        /// </param>
        /// <returns>
        /// True if a ribbon control is visible otherwise false.
        /// </returns>
        bool GetVisible(IRibbonControl control);

        /// <summary>
        /// Callback that returns true if the control is visible. By default this method should return true
        /// </summary>
        /// <param name="control">
        /// The ribbon control.
        /// </param>
        /// <returns>
        /// True if a ribbon control is visible otherwise false.
        /// </returns>
        bool GetMsoVisible(IRibbonControl control);

        /// <summary>
        /// Callback for a custom image.
        /// </summary>
        /// <param name="control">
        /// The ribbon control.
        /// </param>
        /// <returns>
        /// The ribbon control's image.
        /// </returns>
        Bitmap GetImage(IRibbonControl control);

        /// <summary>
        /// Callback fired on user a toggle button press.
        /// </summary>
        /// <param name="control">
        /// The ribbon control.
        /// </param>
        /// <param name="pressed">
        /// Indicates that whether the toggle button is pressed or not.
        /// </param>
        void OnToggleButtonAction(IRibbonControl control, bool pressed);

        /// <summary>
        /// Callback fired on user a button press
        /// </summary>
        /// <param name="control">
        /// The ribbon control.
        /// </param>
        void OnButtonAction(IRibbonControl control);

        /// <summary>
        /// Callback for whether the button is pressed.
        /// </summary>
        /// <param name="control">
        /// The ribbon control.
        /// </param>
        /// <returns>
        /// True if a ribbon button is pressed otherwise false.
        /// </returns>
        bool GetPressed(IRibbonControl control);

        /// <summary>
        /// Callback that returns true if the control is enabled.
        /// </summary>
        /// <param name="control">
        /// The ribbon control.
        /// </param>
        /// <returns>
        /// True if a ribbon control is enabled otherwise false.
        /// </returns>
        bool GetEnabled(IRibbonControl control);

        /// <summary>
        /// Callback that returns the dynamic content for this control.
        /// </summary>
        /// <param name="control">
        /// The ribbon control.
        /// </param>
        /// <returns>
        /// The ribbon control's content.
        /// </returns>
        string GetContent(IRibbonControl control);

        /// <summary>
        /// Callback that sets the label.
        /// </summary>
        /// <param name="control">
        /// The ribbon control.
        /// </param>
        /// <returns>
        /// The ribbon control's label
        /// </returns>
        string GetLabel(IRibbonControl control);

        /// <summary>
        /// Callback that sets the screen tip, which appears on mouse hover.
        /// </summary>
        /// <param name="control">
        /// The ribbon control.
        /// </param>
        /// <returns>
        /// The screen tip, which appears on mouse hover.
        /// </returns>
        /// <remarks>
        /// See <see cref="http://msdn.microsoft.com/en-us/library/bb387063.aspx"/>
        /// </remarks>
        string GetScreentip(IRibbonControl control);

        /// <summary>
        /// Callback that sets the supertip, a large screentip.
        /// </summary>
        /// <param name="control">
        /// The ribbon control.
        /// </param>
        /// <returns>
        /// The supertip, a large screentip.
        /// </returns>
        /// <remarks>
        /// See <see cref="http://msdn.microsoft.com/en-us/library/bb387063.aspx"/>
        /// </remarks>
        string GetSupertip(IRibbonControl control);

        /// <summary>
        /// A callback to load all images.
        /// If specified, all controls with the image attribute set will call this callback with the attribute value passed as a string.
        /// </summary>
        /// <param name="itemId"></param>
        /// <returns></returns>
        Bitmap LoadImage(string itemId);

        /// <summary>
        /// Callback for the number of items in the dropdown.
        /// </summary>
        /// <param name="control">The ribbon control (dropdown).</param>
        /// <returns>The number of items in the dropdown.</returns>
        int GetItemCount(IRibbonControl control);

        /// <summary>
        /// Callback for a combo box, drop-down list, or gallery to get the ID for a specific item.
        /// </summary>
        /// <param name="control">The ribbon control</param>
        /// <param name="itemIndex">Item index</param>
        /// <returns>The ID for a specific item</returns>
        string GetItemId(IRibbonControl control, int itemIndex);

        /// <summary>
        /// Callback for a combo box, drop-down list, or gallery to gets the label for a specific item.
        /// </summary>
        /// <param name="control">The ribbon control</param>
        /// <param name="itemIndex">Item index</param>
        /// <returns>The label for a specific item.</returns>
        string GetItemLabel(IRibbonControl control, int itemIndex);

        /// <summary>
        /// Callback for a drop-down list or gallery to get the ID of the selected item.
        /// </summary>
        /// <param name="control">The ribbon control</param>
        /// <returns>The ID of the selected item.</returns>
        string GetSelectedItemId(IRibbonControl control);

        /// <summary>
        /// Asks for the item that should be selected by index.
        /// </summary>
        /// <param name="control">The ribbon control</param>
        /// <returns>Selected item index</returns>
        int GetSelectedItemIndex(IRibbonControl control);

        /// <summary>
        /// OnAction reports about changing the selected item(in dropdown or gallery).
        /// </summary>
        /// <param name="control">The ribbon control</param>
        /// <param name="selectedItemId">Selected item id</param>
        /// <param name="selectedItemIndex">Selected item index</param>
        void OnChangeAction(IRibbonControl control, string selectedItemId, int selectedItemIndex);

        /// <summary>
        /// Called when the user commits text in an edit box or combo box.
        /// </summary>
        /// <param name="control">The ribbon control</param>
        /// <param name="selectedItemIndex">Selected item index</param>
        void OnChange(IRibbonControl control, string text);

        /// <summary>
        /// Get mask for image. Needed only for Office 2003 command bar
        /// </summary>
        /// <param name="сontrol">The ribbon control</param>
        /// <returns>Return mask of image</returns>
        //Image GetImageMask(IRibbonControl control);

        /// <summary>
        /// Called when the user commits text in an edit box or combo box.
        /// </summary>
        /// <param name="control">The ribbon control</param>
        /// <param name="selectedItemIndex">Selected item index</param>
        string GetText(IRibbonControl control);


        /// <summary>
        /// Called when user press alt
        /// </summary>
        /// <param name="control">The ribbon control</param>
        /// <returns>The ribbon control's keytip</returns>
        string GetKeytip(IRibbonControl control);

        bool ContainsElement(IRibbonControl control);
    }
}
