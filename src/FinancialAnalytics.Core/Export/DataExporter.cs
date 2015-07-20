using System;
using System.Collections.Generic;
using System.Reflection;
using FinancialAnalytics.Wrappers.Excel;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Core.Export
{
    // Exports data into excel worksheet
    public class DataExporter<T> : IDataExporter<T>
    {
        public void DownInsert(IRange insertCell, string[] items, IList<T> data)
        {
            using (var headers = insertCell.get_Resize(1, items.Length) as Range)
            {
                InsertDataIntoRange(headers, items);
                SetHeaderStyle(headers);
            }

            //Write data
            object[,] objData = new object[data.Count, items.Length];
            Type type = typeof(T);
            for (int i = 0; i < data.Count; i++)
            {
                T item = data[i];
                for (int j = 0; j < items.Length; j++)
                {
                    var y = type.InvokeMember(items[j], BindingFlags.GetProperty, null, item, null);
                    objData[i, j] = (y == null) ? "" : y.ToString();
                }
            }
            using (var cell = insertCell.Offset(1, 0))
            {
                using (var range = cell.get_Resize(data.Count, items.Length))
                {
                    InsertDataIntoRange(range, objData);
                    SetDataStyle(range);
                }
            }
            //AutoFitColumns(insertCell, data.Count + 1, objHeaders.Length);
        }

        public void AcrossInsert(IRange insertCell, string[] items, IList<T> data)
        {
            using (var headers = insertCell.get_Resize(items.Length, 1) as Range)
            {
                string[,] rowHeaders = new string[items.Length, 1];
                for (int i = 0; i < items.Length; i++)
                {
                    rowHeaders[i, 0] = items[i];
                }
                InsertDataIntoRange(headers, rowHeaders);
                SetHeaderStyle(headers);
            }

            //Write data
            object[,] objData = new object[items.Length, data.Count];
            Type type = typeof(T);
            for (int i = 0; i < items.Length; i++)
            {

                for (int j = 0; j < data.Count; j++)
                {
                    T item = data[j];
                    var y = type.InvokeMember(items[i], BindingFlags.GetProperty, null, item, null);
                    objData[i, j] = (y == null) ? "" : y.ToString();
                }
            }
            using (var cell = insertCell.Offset(0, 1))
            {
                using (var range = cell.get_Resize(items.Length, data.Count))
                {
                    InsertDataIntoRange(range, objData);
                    SetDataStyle(range);
                }
            }
            //AutoFitColumns(insertCell, objHeaders.Length, data.Count + 1);
        }

        private void InsertDataIntoRange(IRange range, object values)
        {
            if (range != null)
            {
                range.SetValue(values);
            }
        }

        private void SetHeaderStyle(IRange range)
        {
            var font = range.Font;
            font.Bold = true;
        }

        private void AutoFitColumns(IRange range, int rowCount, int colCount)
        {
            using (var rangeToFit = range.get_Resize(rowCount, colCount))
            {
                using (var columns = rangeToFit.Columns)
                {
                    columns.AutoFit();
                }
            }
        }

        private void SetDataStyle(IRange range)
        {
            if (range != null)
            {
                var font = range.Font;
                font.Bold = false;
            }
        }
    }
}
