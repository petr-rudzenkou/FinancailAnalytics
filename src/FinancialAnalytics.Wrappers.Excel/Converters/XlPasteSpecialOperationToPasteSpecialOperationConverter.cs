using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlPasteSpecialOperationToPasteSpecialOperationConverter
    {
        public static PasteSpecialOperation Convert(XlPasteSpecialOperation xlPasteSpecialOperation)
        {
            PasteSpecialOperation pasteSpecialOperation;
            switch (xlPasteSpecialOperation)
            {
                case XlPasteSpecialOperation.xlPasteSpecialOperationAdd:
                    pasteSpecialOperation = PasteSpecialOperation.PasteSpecialOperationAdd;
                    break;
                case XlPasteSpecialOperation.xlPasteSpecialOperationDivide:
                    pasteSpecialOperation = PasteSpecialOperation.PasteSpecialOperationDivide;
                    break;
                case XlPasteSpecialOperation.xlPasteSpecialOperationMultiply:
                    pasteSpecialOperation = PasteSpecialOperation.PasteSpecialOperationMultiply;
                    break;
                case XlPasteSpecialOperation.xlPasteSpecialOperationSubtract:
                    pasteSpecialOperation = PasteSpecialOperation.PasteSpecialOperationSubtract;
                    break;
                default:
                    pasteSpecialOperation = PasteSpecialOperation.PasteSpecialOperationNone;
                    break;
            }
            return pasteSpecialOperation;
        }

        public static XlPasteSpecialOperation ConvertBack(PasteSpecialOperation pasteSpecialOperation)
        {
            XlPasteSpecialOperation xlPasteSpecialOperation;
            switch (pasteSpecialOperation)
            {
                case PasteSpecialOperation.PasteSpecialOperationAdd:
                    xlPasteSpecialOperation = XlPasteSpecialOperation.xlPasteSpecialOperationAdd;
                    break;
                case PasteSpecialOperation.PasteSpecialOperationDivide:
                    xlPasteSpecialOperation = XlPasteSpecialOperation.xlPasteSpecialOperationDivide;
                    break;
                case PasteSpecialOperation.PasteSpecialOperationMultiply:
                    xlPasteSpecialOperation = XlPasteSpecialOperation.xlPasteSpecialOperationMultiply;
                    break;
                case PasteSpecialOperation.PasteSpecialOperationSubtract:
                    xlPasteSpecialOperation = XlPasteSpecialOperation.xlPasteSpecialOperationSubtract;
                    break;
                default:
                    xlPasteSpecialOperation = XlPasteSpecialOperation.xlPasteSpecialOperationNone;
                    break;
            }
            return xlPasteSpecialOperation;
        }
    }
}
