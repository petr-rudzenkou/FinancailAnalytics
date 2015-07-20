using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Image = System.Windows.Controls.Image;

namespace FinancialAnalytics.Views.Base.Controls
{
    public class RemoveButton : Button
    {
        public RemoveButton()
        {
            Width = 16;
            Height = 16;
            BorderThickness = new Thickness(0);
            Cursor = Cursors.Hand;
            var img = new Image();
            BitmapImage bitmapImage;
            using (var memory = new MemoryStream())
            {
                FinancialAnalytics.Resources.ViewsResources.icon_close.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
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
