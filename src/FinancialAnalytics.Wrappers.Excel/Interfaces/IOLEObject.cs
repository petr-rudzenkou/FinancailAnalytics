using System;  
using System.Runtime.InteropServices;
using FinancialAnalytics.Wrappers.Office.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    [ComVisible(true)]
    public interface IOLEObject : IEntityWrapper<IOLEObject>
    {
        IApplication Application { get; }

        string Name { get; }
		string ProgID { get; }
		object Object { get; }
		object Parent { get; }

        IOLEObject Activate();
        IOLEObject Select(bool replace); 
        void CopyPicture(PictureAppearance appearance, CopyPictureFormat format);
        void Copy();
        IOLEObject Delete();

        double Height { get; set; }
        double Width { get; set; }
		double Top { get; set; }
        double Left { get; set; }
    }
}
