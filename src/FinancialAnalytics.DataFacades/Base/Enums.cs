using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.DataFacades.Base
{
    public enum RequestMethod
    {
        POST,
        GET
    }

    public enum ConnectionState
    {
        /// <summary>
        /// Download process completed successfully without timeout, errors or cancelations
        /// </summary>
        /// <remarks></remarks>
        Success,
        /// <summary>
        /// Download process was canceled by user interaction
        /// </summary>
        /// <remarks></remarks>
        Canceled,
        /// <summary>
        /// Download process reached the setted timeout span
        /// </summary>
        /// <remarks></remarks>
        Timeout,
        /// <summary>
        /// An Error occured during download process
        /// </summary>
        /// <remarks></remarks>
        ErrorOccured
    }
}
