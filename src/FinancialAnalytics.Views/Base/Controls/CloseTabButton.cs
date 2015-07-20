using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace FinancialAnalytics.Views.Base.Controls
{
    public class CloseTabButton : Button
    {
        public CloseTabButton()
        {
            Width = 12;
            Height = 12;
            BorderThickness = new Thickness(0);
            Cursor = Cursors.Hand;
            Margin = new Thickness(5, 0, 0, 0);
            var img = new Image();
            BitmapImage bitmapImage;
            using (var memory = new MemoryStream())
            {
                FinancialAnalytics.Resources.ViewsResources.closeTabButton.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                memory.Position = 0;
                bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
            }

            img.Source = bitmapImage;
            Content = img;
        }
    }
}
