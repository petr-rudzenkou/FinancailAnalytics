using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;

namespace FinancialAnalytics.ExcelUI
{
    public static class IconManager
    {
        public static Bitmap GetIcon16(IEnumerable<Bitmap> images)
        {
            return GetIcon(images, 16);
        }

        public static Bitmap GetIcon32(IEnumerable<Bitmap> images)
        {
            return GetIcon(images, 32);
        }

        private static Bitmap GetIcon(IEnumerable<Bitmap> images, int width)
        {
            if (images == null)
                return null;
            var imagesArray = images.ToArray();
            if (imagesArray.Length == 0)
                return null;
            var larger = imagesArray.OrderBy(i => i.Width).FirstOrDefault(i => i.Width >= width);
            var smaller = imagesArray.OrderByDescending(i => i.Width).FirstOrDefault(i => i.Width <= width);
            return larger ?? smaller;
        }

        public static Bitmap CreateIcon(Image image, Brush backgroundBrush)
        {
            var backgrounded = new Bitmap(image.Width, image.Height);
            using (var g = Graphics.FromImage(backgrounded))
            {
                g.FillRectangle(backgroundBrush, 0, 0, image.Width, image.Height);
                g.DrawImage(image, 0, 0);
            }
            return backgrounded;
        }

        public static Bitmap CreateMask(Bitmap image, Func<Color, bool> isTransparent)
        {
            Bitmap mask = new Bitmap(image.Width, image.Height);
            for (var x = 0; x < image.Width; x++)
                for (var y = 0; y < image.Height; y++)
                    mask.SetPixel(x, y, (isTransparent(image.GetPixel(x, y)) ? Color.White : Color.Black));
            return mask;
        }

        /// <summary>
        /// Eliminates resizing distortions at the borders
        /// </summary>
        public static Bitmap ResizeIcon(Image image, int newWidth, int newHeight)
        {
            // draw large borders to cut border distortions later
            using (var largeBackgrounded = new Bitmap(image.Width * 3, image.Height * 3))
            {
                using (var g = Graphics.FromImage(largeBackgrounded))
                {
                    g.DrawImage(image, image.Width, image.Height);
                }

                // resize
                using (var resizedLarge = ResizeImage(largeBackgrounded, newWidth * 3, newHeight * 3))
                {
                    // crop to original	size
                    var cropped = resizedLarge.Clone(new Rectangle(newWidth, newHeight, newWidth, newHeight), image.PixelFormat);

                    return cropped;
                }
            }
        }

        private static Bitmap ResizeImage(Image image, int newWidth, int newHeight)
        {
            Bitmap newImage = new Bitmap(newWidth, newHeight);
            using (var gr = Graphics.FromImage(newImage))
            {
                gr.SmoothingMode = SmoothingMode.AntiAlias;
                gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gr.PixelOffsetMode = PixelOffsetMode.HighQuality;
                gr.DrawImage(image, new Rectangle(0, 0, newWidth, newHeight));
            }
            return newImage;
        }
    }
}
