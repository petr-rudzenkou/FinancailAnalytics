﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.DataFacades.Base
{
    public class ConversionException : Exception
    {
        public ConversionException() { }
        public ConversionException(string message) : base(message) { }
        public ConversionException(string message, Exception innerException) : base(message, innerException) { }
    }
}
