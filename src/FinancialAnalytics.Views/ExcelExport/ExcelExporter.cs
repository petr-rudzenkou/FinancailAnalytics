using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using FinancialAnalytics.Core;
using FinancialAnalytics.Core.Export;
using FinancialAnalytics.Wrappers.Excel;

namespace FinancialAnalytics.Views.ExcelExport
{
    public class ExcelExporter : IExcelExporter
    {
        private const string FORMULA_SEPATOR = ";";

        private readonly IApplicationProvider _applicationProvider;
        private readonly IDataExporterFactory _dataExporterFactory;
        public ExcelExporter(IApplicationProvider applicationProvider, IDataExporterFactory dataExporterFactory)
        {
            _applicationProvider = applicationProvider;
            _dataExporterFactory = dataExporterFactory;
        }

        public void InsertTickets(IEnumerable<string> tickets)
        {
            string formula = ConstructFormula(tickets);

            Task.Factory.StartNew(() =>
            {
                _applicationProvider.WhenReady(x =>
                {
                    using (var activeCell = x.ActiveCell)
                    {
                        activeCell.Formula = formula;
                    }
                });
            });

        }

        public void InsertTickets(IEnumerable<string> tickets, string layout)
        {
            throw new NotImplementedException();
        }

        public void InsertTickets(IEnumerable<string> tickets, IEnumerable<string> dataItems)
        {
            throw new NotImplementedException();
        }

        public void InsertTickets(IEnumerable<string> tickets, IEnumerable<string> dataItems, string layout)
        {
            throw new NotImplementedException();
        }

        public void InsertData<T>(IEnumerable<T> data, string[] headers = null)
        {
            Type type = typeof(T);
            var properties = type.GetProperties().Select(x => x.Name).ToList();
            if (headers != null)
            {
                properties.RemoveAll(x => !headers.Contains(x));
            }
            var dataExproter = _dataExporterFactory.Create<T>();
            using (var activeCell = _applicationProvider.Application.ActiveCell)
            {
                dataExproter.DownInsert(activeCell, properties.ToArray(), data.ToList());
            }
        }

        private string ConstructFormula(IEnumerable<string> tickets, IEnumerable<string> dataItems = null, string layout = null)
        {
            var ticketsArray = tickets.ToArray();
            var formula = new StringBuilder();
            formula.Append("=FA(");
            formula.Append('"');
            for (int i = 0; i < ticketsArray.Length; i++)
            {
                formula.Append(ticketsArray[i]);
                if (i != ticketsArray.Length - 1)
                {
                    formula.Append(",");
                }
                else
                {
                    formula.Append('"'); ;
                }
            }

            if (dataItems != null)
            {
                formula.Append(FORMULA_SEPATOR);
                formula.Append('"');
                var dataItemsArray = dataItems.ToArray();
                for (int i = 0; i < dataItemsArray.Length; i++)
                {
                    formula.Append(dataItemsArray[i]);
                    if (i != dataItemsArray.Length - 1)
                    {
                        formula.Append(",");
                    }
                    else
                    {
                        formula.Append('"');
                    }
                }
            }

            if (!string.IsNullOrEmpty(layout))
            {
                if (dataItems == null)
                {
                    formula.Append(FORMULA_SEPATOR);
                }
                formula.Append(FORMULA_SEPATOR);
                formula.Append('"');
                formula.Append(layout);
                formula.Append('"');
            }

            formula.Append(")");
            if (formula.Length > 256)
            {
                MessageBox.Show("Too long formula. Maximum size is 256 characters. Please, specify fewer tickers.");
            }
            return formula.ToString();
        }

        public void InsertImage(BitmapImage image)
        {
            Task.Factory.StartNew(() =>
            {
                _applicationProvider.WhenReady(x =>
                {
                    using (var cell = x.ActiveCell)
                    {
                        
                        using (Worksheet worksheet = cell.Worksheet as Worksheet)
                        {
                            if (worksheet != null)
                            {
                                //TODO: Temperary solution.
                                string currentBuffer = Clipboard.GetText();
                                Clipboard.SetImage(image);
                                worksheet.Paste();
                                if (!string.IsNullOrEmpty(currentBuffer))
                                {
                                    Clipboard.SetText(currentBuffer);
                                }
                            }
                        }
                    }

                });
            });
        }
    }
}
