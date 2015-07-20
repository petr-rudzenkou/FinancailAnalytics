using System.Windows;

namespace FinancialAnalytics.Views.ViewSettings
{
    public class WindowSettings
    {
        public double? Top { get; set; }
     
        public double? Left { get; set; }
 
        public double? Height { get; set; }
        
        public double? Width { get; set; }
        
        public bool? TopMost { get; set; }
        
        public string Title { get; set; }
   
        public long Parent { get; set; }
        
        public WindowStyle? Style { get; set; }

        public ResizeMode? ResizeMode { get; set; }
        
        public bool? IsMaximizeButtonVisible { get; set; }
        
        public bool? IsMinimizeButtonVisible { get; set; }
        
        public string ChildOf { get; set; }
        
        public string HelpTopicId { get; set; }
        
        public string Id { get; set; }
        
        public bool? ApplyScaling { get; set; }
    }
}
