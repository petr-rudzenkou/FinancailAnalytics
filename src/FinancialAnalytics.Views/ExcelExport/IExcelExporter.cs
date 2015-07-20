using System.Collections.Generic;
using System.Windows.Media.Imaging;

namespace FinancialAnalytics.Views.ExcelExport
{
    public interface IExcelExporter
    {
        void InsertTickets(IEnumerable<string> tickets);
        void InsertTickets(IEnumerable<string> tickets, string layout);
        void InsertTickets(IEnumerable<string> tickets, IEnumerable<string> dataItems);
        void InsertTickets(IEnumerable<string> tickets, IEnumerable<string> dataItems, string layout);

        void InsertData<T>(IEnumerable<T> data, string[] headers = null);

        void InsertImage(BitmapImage image);
    }
}
