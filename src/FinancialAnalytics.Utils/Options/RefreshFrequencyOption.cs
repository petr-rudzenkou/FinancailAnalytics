namespace FinancialAnalytics.Utils.Options
{
    public class RefreshFrequencyOption : OptionBase
    {
        public RefreshFrequencyOption()
        {
            Measures = new string[]
            {
                Options.RefreshFrequencyMeasure.Sec,
                Options.RefreshFrequencyMeasure.Min,
                Options.RefreshFrequencyMeasure.Hours,
                Options.RefreshFrequencyMeasure.Days
            };
        }

        public int? RefreshFrequency { get; set; }
        public string RefreshFrequencyMeasure { get; set; }
        public string[] Measures { get; set; }
    }
}
