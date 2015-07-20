using System;
using System.Globalization;

namespace FinancialAnalytics.Wrappers.Excel
{
    public class WorkbookTemplateFactory
    {
        public static object GetDefaultTemplate(CultureInfo cultureInfo)
        {
            switch (cultureInfo.TwoLetterISOLanguageName)
            {
                case "fr":
                    return "Classeur";
                case "ja":
                    return "ブック";
                case "ko":
                    return "통합 문서";
                case "ru":
                    return "Книга";
                case "zh":
                    return "工作簿";
                case "de":
                    return "Arbeitsmappe";
                case "en":
                    return "Workbook";
                default:
                    return Type.Missing;
            } 
        }
    }
}
