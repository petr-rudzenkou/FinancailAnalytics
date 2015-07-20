using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.DataFacades.Interfaces
{
    /// <summary>
    /// Provides properties to set the start index and count number for a query in results queue.
    /// </summary>
    public interface IResultIndexSettings
    {
        /// <summary>
        /// The results queue start index.
        /// </summary>
        int Index { get; set; }
        /// <summary>
        /// The total number of results.
        /// </summary>
        int Count { get; set; }
    }
}
