using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.DataFacades.Base;

namespace FinancialAnalytics.DataFacades.Charts
{
    public class ChartDownloadSettings : SettingsBase
    {
        private string mID = string.Empty;
        private Culture mCulture = null;
        private int mImageWidth = 300;
        private int mImageHeight = 180;
        private bool mSimplifiedImage = false;
        private ChartImageSize mImageSize = ChartImageSize.Middle;
        private ChartTimeSpan mTimeSpan = ChartTimeSpan.c1D;
        private ChartType mType = ChartType.Line;
        private ChartScale mScale = ChartScale.Logarithmic;
        private List<MovingAverageInterval> mMovingAverages = new List<MovingAverageInterval>();
        private List<MovingAverageInterval> mEMovingAverages = new List<MovingAverageInterval>();
        private List<TechnicalIndicator> mTechnicalIndicators = new List<TechnicalIndicator>();
        private List<ChartOverlay> mChartOverlays = new List<ChartOverlay>();
        private List<string> mComparingIDs = new List<string>();

        public string ID
        {
            get
            {
                return mID;
            }
            set
            {
                mID = value;
            }
        }
        /// <summary>
        /// Gets or sets the used culture for scale descriptions. Can only be used with Server [USA].
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public Culture Culture
        {
            get { return mCulture; }
            set { mCulture = value; }
        }
        /// <summary>
        /// Gets the width of the image.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int ImageWidth
        {
            get { return mImageWidth; }
            set { mImageWidth = value; }
        }
        /// <summary>
        /// Gets the height of the image
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int ImageHeight
        {
            get { return mImageHeight; }
            set { mImageHeight = value; }
        }
        /// <summary>
        /// Gets a bool value if the image is simplified (1 day period; only ImageWidth, ImageHeight and Culture options available; Other options will be ignored)
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool SimplifiedImage
        {
            get { return mSimplifiedImage; }
            set { mSimplifiedImage = value; }
        }
        /// <summary>
        /// Gets the size of the image (only available if SimplifiedImage = FALSE)
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ChartImageSize ImageSize
        {
            get { return mImageSize; }
            set { mImageSize = value; }
        }
        /// <summary>
        /// Gets the span of the reviewed period.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ChartTimeSpan TimeSpan
        {
            get { return mTimeSpan; }
            set { mTimeSpan = value; }
        }
        /// <summary>
        /// Gets the chart type of the image.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ChartType Type
        {
            get { return mType; }
            set { mType = value; }
        }
        /// <summary>
        /// Gets the scaling of the image.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ChartScale Scale
        {
            get { return mScale; }
            set { mScale = value; }
        }
        /// <summary>
        /// Gets the list of moving average indicators.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public List<MovingAverageInterval> MovingAverages
        {
            get { return mMovingAverages; }
            set { mMovingAverages = value; }
        }
        /// <summary>
        /// Gets the list of exponential moving average indicators.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public List<MovingAverageInterval> ExponentialMovingAverages
        {
            get { return mEMovingAverages; }
            set { mEMovingAverages = value; }
        }
        /// <summary>
        /// Gets the list of technical indicators.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public List<TechnicalIndicator> TechnicalIndicators
        {
            get { return mTechnicalIndicators; }
            set { mTechnicalIndicators = value; }
        }

        public List<ChartOverlay> ChartOverlays
        {
            get { return mChartOverlays; }
            set { mChartOverlays = value; }
        }
        /// <summary>
        /// Gets the ID list of all compared stocks/indices.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public List<string> ComparingIDs
        {
            get { return mComparingIDs; }
            set { mComparingIDs = value; }
        }

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <remarks></remarks>
        public ChartDownloadSettings()
        {
            mCulture = Culture.DefaultCultures.UnitedStates_English;
        }
        public ChartDownloadSettings(string id)
        {
            mCulture = Culture.DefaultCultures.UnitedStates_English;
            this.ID = id;
        }
        public override string GetUrl()
        {
            if (ID == string.Empty) { throw new ArgumentException("ID is empty.", "ID"); }
            StringBuilder url = new StringBuilder();
            url.Append("http://");
            url.Append("chart.finance.yahoo.com/");

            if (this.SimplifiedImage) { url.Append("t?"); }
            else if (this.ImageSize == ChartImageSize.Small) { url.Append("h?"); }
            else { url.Append("z?"); }

            url.Append("s=");
            url.Append(DataFacadesHelper.CleanYqlParam(FinanceHelper.CleanIndexID(this.ID)));

            if (this.SimplifiedImage)
            {
                url.Append("&width=" + this.ImageWidth.ToString());
                url.Append("&height=" + this.ImageHeight.ToString());
            }
            else if (this.ImageSize != ChartImageSize.Small)
            {
                url.Append("&t=");
                url.Append(FinanceHelper.GetChartTimeSpan(this.TimeSpan));
                url.Append("&z=");
                url.Append(FinanceHelper.GetChartImageSize(this.ImageSize));
                url.Append("&q=");
                url.Append(FinanceHelper.GetChartType(this.Type));
                url.Append("&l=");
                url.Append(FinanceHelper.GetChartScale(this.Scale));
                if (this.MovingAverages.Count > 0 | this.ExponentialMovingAverages.Count > 0 | this.TechnicalIndicators.Count > 0 || ChartOverlays.Count > 0)
                {
                    url.Append("&p=");
                    foreach (MovingAverageInterval ma in this.MovingAverages)
                    {
                        url.Append('m');
                        url.Append(FinanceHelper.GetMovingAverageInterval(ma));
                        url.Append(',');
                    }
                    foreach (MovingAverageInterval ma in this.ExponentialMovingAverages)
                    {
                        url.Append('e');
                        url.Append(FinanceHelper.GetMovingAverageInterval(ma));
                        url.Append(',');
                    }
                    foreach (TechnicalIndicator ti in this.TechnicalIndicators)
                    {
                        url.Append(FinanceHelper.GetTechnicalIndicatorsI(ti));
                    }
                    for(int i = 0; i < ChartOverlays.Count; i++)
                    {
                        url.Append(ChartOverlays[i]);
                        if (i < ChartOverlays.Count - 1)
                        {
                            url.Append(',');
                        }
                    }
                }
                if (this.TechnicalIndicators.Count > 0)
                {
                    url.Append("&a=");
                    foreach (TechnicalIndicator ti in this.TechnicalIndicators)
                    {
                        url.Append(FinanceHelper.GetTechnicalIndicatorsII(ti));
                    }
                }
                if (this.ComparingIDs.Count > 0)
                {
                    url.Append("&c=");
                    foreach (string csid in this.ComparingIDs)
                    {
                        url.Append(DataFacadesHelper.CleanYqlParam(FinanceHelper.CleanIndexID(csid)));
                        url.Append(',');
                    }
                }
                if (this.Culture == null)
                {
                    this.Culture = Culture.DefaultCultures.UnitedStates_English;
                }
            }
            url.Append(DataFacadesHelper.CultureToParameters(this.Culture));
            return url.ToString();
        }

        public override object Clone()
        {
            ChartDownloadSettings cln = new ChartDownloadSettings();
            cln.ID = this.ID;
            cln.SimplifiedImage = this.SimplifiedImage;
            cln.ImageWidth = this.ImageWidth;
            cln.ImageHeight = this.ImageHeight;
            cln.ImageSize = this.ImageSize;
            cln.TimeSpan = this.TimeSpan;
            cln.Type = this.Type;
            cln.Scale = this.Scale;
            cln.Culture = this.Culture;
            cln.MovingAverages.AddRange((MovingAverageInterval[])this.MovingAverages.ToArray().Clone());
            cln.ExponentialMovingAverages.AddRange((MovingAverageInterval[])this.ExponentialMovingAverages.ToArray().Clone());
            cln.TechnicalIndicators.AddRange((TechnicalIndicator[])this.TechnicalIndicators.ToArray().Clone());
            cln.ComparingIDs.AddRange((string[])this.ComparingIDs.ToArray().Clone());
            return cln;
        }
    }
}
