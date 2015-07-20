using System.Drawing;
using System.Windows.Forms;
using stdole;

namespace FinancialAnalytics.Wrappers.Office
{
	public static class PictureDispConverter
	{
		// ReSharper disable ClassNeverInstantiated.Local
		private class AxHostConverter : AxHost
		// ReSharper restore ClassNeverInstantiated.Local
		{
			private AxHostConverter() : base("") { }

			public static IPictureDisp ImageToPictureDisp(Image image)
			{
				return (IPictureDisp)GetIPictureDispFromPicture(image);
			}

			public static Image PictureDispToImage(IPictureDisp pictureDisp)
			{
				return GetPictureFromIPicture(pictureDisp);
			}
		}

		public static IPictureDisp Convert(Bitmap image)
		{
			return AxHostConverter.ImageToPictureDisp(image);
		}

		public static Bitmap Convert(IPictureDisp pictureDisp)
		{
			return (Bitmap)AxHostConverter.PictureDispToImage(pictureDisp);
		}
	}
}