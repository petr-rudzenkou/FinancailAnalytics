using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ISheet : IEntityWrapper<ISheet>
    {
        string Name { get; set; }

        object SheetObject { get; }

        SheetVisibility Visible { get; set; }

        string CodeName { get; }
        
        void Activate();

        void Delete();

        void Select(bool replace);

		IWorkbook Workbook { get; }

        void Calculate();
    }
}
