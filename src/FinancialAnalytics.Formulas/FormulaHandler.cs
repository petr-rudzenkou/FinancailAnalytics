using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using FinancialAnalytics.Core;
using FinancialAnalytics.Core.Export;
using FinancialAnalytics.DataFacades.Quotes;
using FinancialAnalytics.DataFacades.XChangeRates;
using FinancialAnalytics.Formulas.Formulas;
using FinancialAnalytics.Wrappers.Excel;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Formulas
{
    public class FormulaHandler : IFormulaHandler
    {
        private readonly IApplicationProvider _applicationProvider;
        private readonly QuotesDownload _quotesDownload;
        private readonly IDataExporter<QuotesData> _dataExporter;

        public FormulaHandler(IApplicationProvider applicationProvider, IDataExporterFactory dataExporterFactory)
        {
            _applicationProvider = applicationProvider;
            _quotesDownload = new QuotesDownload();
            _dataExporter = dataExporterFactory.Create<QuotesData>();
        }

        public object FA(object symbols, [OptionalAttribute]object dataItems, [OptionalAttribute]object layout, [OptionalAttribute]object destinationCell)
        {
            IRange callerCell = _applicationProvider.Application.GetCaller();

            string[] arraySymbols = ExtractSymbols(symbols);
            string[] items = ExstractDataItems(dataItems);
            string definedLayout = ExtractLayout(layout);
            string address = callerCell.Address;
            IRange destCell = ExstractDestinationCell(destinationCell);

            if (!FormulasRegistry.Contains(address))
            {
                FormulasRegistry.Register(address, new FormulaItem
                {
                    Symbols = arraySymbols,
                    DataItems = items,
                    Layout = definedLayout,
                });
            }

            FormulaItem formulaItem = FormulasRegistry.GetFormulaItem(address);

            if (formulaItem.WithDestinationCell)
            {
                formulaItem.WithDestinationCell = false;
                return "Retrieving...";
            }
            if (formulaItem.Handled)
            {
                formulaItem.Handled = false;
                callerCell.Dispose();
                return "Updated: " + DateTime.Now;
            }
            if (formulaItem.Error)
            {
                formulaItem.Error = false;
                formulaItem.Handled = false;
                callerCell.Dispose();
                return "Error";
            }
            formulaItem.WithDestinationCell = destCell != null;

            Task.Factory.StartNew(() =>
            {
                var quotes = _quotesDownload.Download(arraySymbols).Result.Items;
                _applicationProvider.WhenReady(x =>
                {
                    if (destCell == null)
                    {
                        destCell = callerCell.Offset(1, 0);
                    }
                    if (ExtractLayout(definedLayout).Equals("Across", StringComparison.InvariantCultureIgnoreCase))
                    {
                        _dataExporter.AcrossInsert(destCell, items, quotes);
                    }
                    else
                    {
                        _dataExporter.DownInsert(destCell, items, quotes);
                    }
                    destCell.Dispose();

                    formulaItem.Handled = true;
                    string formula = callerCell.Formula;
                    callerCell.Formula = formula;
                    callerCell.Dispose();
                });
            });

            return "Retrieving...";
        }

        private string ExtractLayout(object layout)
        {
            if (layout != null && layout.GetType() != typeof(Missing))
            {
                return layout.ToString();
            }
            return "Down";
        }

        private string[] ExtractSymbols(object symbols)
        {
            var extractedSymbols = new List<string>();
            var symbolsRange = symbols as Microsoft.Office.Interop.Excel.Range;
            if (symbolsRange != null)
            {
                var cells = symbolsRange.Cells;
                foreach (Microsoft.Office.Interop.Excel.Range cell in cells)
                {
                    string text = cell.Text.Trim();
                    if (!string.IsNullOrEmpty(text))
                    {
                        extractedSymbols.AddRange(text.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries));
                    }
                    ComObjectsFinalizer.ReleaseComObject(cell);
                }
                ComObjectsFinalizer.ReleaseComObject(cells);
                ComObjectsFinalizer.ReleaseComObject(symbolsRange);
            }
            else
            {
                var symbolsString = symbols.ToString().Trim();
                var arraySymbols = symbolsString.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                extractedSymbols.AddRange(arraySymbols.Select(x => x.Trim()));
            }
            return extractedSymbols.ToArray();
        }

        private string[] ExstractDataItems(object dataItems)
        {
            PropertyInfo[] headerInfo = typeof(QuotesData).GetProperties();

            var objHeaders = headerInfo.Select(y => y.Name).ToArray();
            
            if (dataItems != null && dataItems.GetType() != typeof(Missing))
            {
                var extractedDataItems = new List<string>();
                var dataItemsRange = dataItems as Microsoft.Office.Interop.Excel.Range;
                if (dataItemsRange != null)
                {
                    var cells = dataItemsRange.Cells;
                    foreach (Microsoft.Office.Interop.Excel.Range cell in cells)
                    {
                        string text = cell.Text.Trim();
                        if (!string.IsNullOrEmpty(text))
                        {
                            var item = objHeaders.FirstOrDefault(x => x.Equals(text, StringComparison.InvariantCultureIgnoreCase));
                            if (item != null)
                            {
                                extractedDataItems.Add(text);
                            }
                        }
                        ComObjectsFinalizer.ReleaseComObject(cell);
                    }
                    ComObjectsFinalizer.ReleaseComObject(cells);
                    ComObjectsFinalizer.ReleaseComObject(dataItemsRange);
                }
                else
                {
                    if (!string.IsNullOrEmpty(dataItems.ToString()))
                    {
                        string[] userDefinedDataItems = dataItems.ToString().Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                        userDefinedDataItems = userDefinedDataItems.Select(y => y.Trim()).ToArray();

                        var intercectedHeaders = new List<string>();
                        for (int i = 0; i < userDefinedDataItems.Length; i++)
                        {
                            for (int j = 0; j < objHeaders.Count(); j++)
                            {
                                if (objHeaders[j].Equals(userDefinedDataItems[i], StringComparison.InvariantCultureIgnoreCase))
                                {
                                    intercectedHeaders.Add(objHeaders[j]);
                                }
                            }
                        }
                        extractedDataItems = intercectedHeaders;
                    }
                }
                return extractedDataItems.ToArray();
            }
            return QuotesProperties.DefaultQuoteProperties;
        }

        private string[] ExstractXChangeRatesDataItems(object dataItems)
        {
            PropertyInfo[] headerInfo = typeof(QuotesData).GetProperties();

            var objHeaders = headerInfo.Select(y => y.Name).ToArray();

            if (dataItems != null && dataItems.GetType() != typeof(Missing))
            {
                var extractedDataItems = new List<string>();
                var dataItemsRange = dataItems as Microsoft.Office.Interop.Excel.Range;
                if (dataItemsRange != null)
                {
                    var cells = dataItemsRange.Cells;
                    foreach (Microsoft.Office.Interop.Excel.Range cell in cells)
                    {
                        string text = cell.Text.Trim();
                        if (!string.IsNullOrEmpty(text))
                        {
                            var item = objHeaders.FirstOrDefault(x => x.Equals(text, StringComparison.InvariantCultureIgnoreCase));
                            if (item != null)
                            {
                                extractedDataItems.Add(text);
                            }
                        }
                        ComObjectsFinalizer.ReleaseComObject(cell);
                    }
                    ComObjectsFinalizer.ReleaseComObject(cells);
                    ComObjectsFinalizer.ReleaseComObject(dataItemsRange);
                }
                else
                {
                    if (!string.IsNullOrEmpty(dataItems.ToString()))
                    {
                        string[] userDefinedDataItems = dataItems.ToString().Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                        userDefinedDataItems = userDefinedDataItems.Select(y => y.Trim()).ToArray();

                        var intercectedHeaders = new List<string>();
                        for (int i = 0; i < userDefinedDataItems.Length; i++)
                        {
                            for (int j = 0; j < objHeaders.Count(); j++)
                            {
                                if (objHeaders[j].Equals(userDefinedDataItems[i], StringComparison.InvariantCultureIgnoreCase))
                                {
                                    intercectedHeaders.Add(objHeaders[j]);
                                }
                            }
                        }
                        extractedDataItems = intercectedHeaders;
                    }
                }
                return extractedDataItems.ToArray();
            }
            return XChangeRatesProperties.DefaultXChangeRatesProperties;
        }

        private IRange ExstractDestinationCell(object destinationCell)
        {
            var cell = destinationCell as Microsoft.Office.Interop.Excel.Range;
            if (cell != null)
            {
                using (var sheet = _applicationProvider.Application.ActiveSheet as Worksheet)
                {
                    string address = cell.Address;
                    ComObjectsFinalizer.ReleaseComObject(cell);
                    return sheet.GetRange(address);
                }
            }
            return null;
        }


        public object FACurrency(object symbols, [OptionalAttribute]object dataItems, [OptionalAttribute]object layout, [OptionalAttribute]object destinationCell)
        {
            IRange callerCell = _applicationProvider.Application.GetCaller();

            string[] arraySymbols = ExtractSymbols(symbols);
            string[] items = ExstractXChangeRatesDataItems(dataItems);
            string definedLayout = ExtractLayout(layout);
            string address = callerCell.Address;
            IRange destCell = ExstractDestinationCell(destinationCell);

            if (!FormulasRegistry.Contains(address))
            {
                FormulasRegistry.Register(address, new FormulaItem
                {
                    Symbols = arraySymbols,
                    DataItems = items,
                    Layout = definedLayout,
                });
            }

            FormulaItem formulaItem = FormulasRegistry.GetFormulaItem(address);

            if (formulaItem.WithDestinationCell)
            {
                formulaItem.WithDestinationCell = false;
                return "Retrieving...";
            }
            if (formulaItem.Handled)
            {
                formulaItem.Handled = false;
                callerCell.Dispose();
                return "Updated: " + DateTime.Now;
            }
            if (formulaItem.Error)
            {
                formulaItem.Error = false;
                formulaItem.Handled = false;
                callerCell.Dispose();
                return "Error";
            }
            formulaItem.WithDestinationCell = destCell != null;

            Task.Factory.StartNew(() =>
            {
                var quotes = _quotesDownload.Download(arraySymbols).Result.Items;
                _applicationProvider.WhenReady(x =>
                {
                    if (destCell == null)
                    {
                        destCell = callerCell.Offset(1, 0);
                    }
                    if (ExtractLayout(definedLayout).Equals("Across", StringComparison.InvariantCultureIgnoreCase))
                    {
                        _dataExporter.AcrossInsert(destCell, items, quotes);
                    }
                    else
                    {
                        _dataExporter.DownInsert(destCell, items, quotes);
                    }
                    destCell.Dispose();

                    formulaItem.Handled = true;
                    string formula = callerCell.Formula;
                    callerCell.Formula = formula;
                    callerCell.Dispose();
                });
            });

            return "Retrieving...";
        }
    }
}
