using System;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Media.Imaging;
using FinancialAnalytics.DataFacades;
using FinancialAnalytics.DataFacades.Base;
using FinancialAnalytics.DataFacades.Charts;
using FinancialAnalytics.Views.Portfolio.Base;

namespace FinancialAnalytics.Views.Portfolio.Converters
{
    public class SmallImageConverter : IValueConverter
    {
        private const string IMAGE_URL_FORMAT_STRING = @"https://chart.finance.yahoo.com/h?s={0}&lang=en-US&region=US";

        public object Convert(object value, Type targetType,
                          object parameter, CultureInfo culture)
        {

            //var task = new Task<BitmapImage>(() =>
            //{
                string symbol = value as string;
                if (symbol != null)
                {
                    //var settings = new ChartDownloadSettings();
                    //settings.ImageSize = ChartImageSize.Small;
                    //var dl = new ChartDownload()
                    //{
                    //    Settings = settings
                    //};
                    //Response<ChartResult> response = dl.Download(symbol);
                    //if (response.Connection.State == ConnectionState.Success)
                    //{
                    //    System.IO.MemoryStream stream = response.Result.Item;
                    //    BitmapImage image = new BitmapImage();
                    //    image.BeginInit();
                    //    image.StreamSource = stream;
                    //    image.EndInit();
                    //    return image;
                    //}
                    var uri = new Uri(string.Format(IMAGE_URL_FORMAT_STRING, symbol));
                    var image = new BitmapImage(uri);
                    return image;
                }
                return new BitmapImage();
            //});
            //task.Start();

            //return new TaskCompletionNotifier<BitmapImage>(task);
        }

        public object ConvertBack(object value, Type targetType,
                                  object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
