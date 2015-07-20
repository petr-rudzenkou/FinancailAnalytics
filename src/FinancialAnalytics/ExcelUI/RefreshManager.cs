using FinancialAnalytics.Core;
using FinancialAnalytics.Wrappers.Excel;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.ExcelUI
{
    public class RefreshManager : IRefreshManager
    {
        private readonly IApplicationProvider _applicationProvider;

        public RefreshManager(IApplicationProvider applicationProvider)
        {
            _applicationProvider = applicationProvider;
        }

        public void Refresh(string refreshMode)
        {
            switch (refreshMode)
            {
                case RibbonIds.FA_REFRESH_ALL_WORKBOOKS:
                    {
                        RefreshAllWorkbooks();
                        break;
                    }
                case RibbonIds.FA_REFRESH_ACTIVE_WORKBOOK:
                    {
                        RefreshActiveWorkbook();
                        break;
                    }
                case RibbonIds.FA_REFRESH_ACTIVE_WORSHEET:
                    {
                        RefreshActiveWorksheet();
                        break;
                    }
                case RibbonIds.FA_REFRESH_ACTIVE_CELL:
                {
                    RefreshSelection();
                        break;
                    }
            }
        }

        private void RefreshAllWorkbooks()
        {
            _applicationProvider.WhenReady(x =>
            {
                using (var workbooks = x.Workbooks)
                {
                    foreach (var workbook in workbooks)
                    {
                        using (workbook)
                        {
                            using (var worksheets = workbook.Worksheets)
                            {
                                foreach (var worksheet in worksheets)
                                {
                                    using (worksheet)
                                    {
                                        using (var usedRange = worksheet.UsedRange)
                                        {
                                            using (var cells = usedRange.Cells)
                                            {
                                                foreach (var cell in cells)
                                                {
                                                    using (cell)
                                                    {
                                                        if (cell.HasFormula.HasValue && cell.HasFormula.Value)
                                                        {
                                                            string formula = cell.Formula;
                                                            cell.Formula = formula;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            });
        }

        private void RefreshActiveWorkbook()
        {
            _applicationProvider.WhenReady(x =>
            {
                using (var workbook = x.ActiveWorkbook)
                {
                    using (var worksheets = workbook.Worksheets)
                    {
                        foreach (var worksheet in worksheets)
                        {
                            using (worksheet)
                            {
                                using (var usedRange = worksheet.UsedRange)
                                {
                                    using (var cells = usedRange.Cells)
                                    {
                                        foreach (var cell in cells)
                                        {
                                            using (cell)
                                            {
                                                if (cell.HasFormula.HasValue && cell.HasFormula.Value)
                                                {
                                                    string formula = cell.Formula;
                                                    cell.Formula = formula;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                        }
                    }
                }
            });
        }

        private void RefreshActiveWorksheet()
        {
            _applicationProvider.WhenReady(x =>
            {
                using (var worksheet = x.ActiveSheet as Worksheet)
                {
                    using (var usedRange = worksheet.UsedRange)
                    {
                        using (var cells = usedRange.Cells)
                        {
                            foreach (var cell in cells)
                            {
                                using (cell)
                                {
                                    if (cell.HasFormula.HasValue && cell.HasFormula.Value)
                                    {
                                        string formula = cell.Formula;
                                        cell.Formula = formula;
                                    }
                                }
                            }
                        }
                    }
                }
            });
        }

        private void RefreshSelection()
        {
            _applicationProvider.WhenReady(x =>
            {
                using (var activeCell = x.Selection as IRange)
                {
                    using (var cells = activeCell)
                    {
                        foreach (var cell in cells)
                        {
                            using (cell)
                            {
                                if (cell.HasFormula.HasValue && cell.HasFormula.Value)
                                {
                                    string formula = cell.Formula;
                                    cell.Formula = formula;
                                }
                            }
                        }
                    }
                }
            });
        }
    }
}
