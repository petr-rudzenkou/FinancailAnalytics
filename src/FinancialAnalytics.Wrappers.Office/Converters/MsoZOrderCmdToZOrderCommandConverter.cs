using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{

    public class MsoZOrderCmdToZOrderCommandConverter
    {
        public static ZOrderCommand Convert(MsoZOrderCmd msoZOrderCmd)
        {
            ZOrderCommand zOrderCommand = ZOrderCommand.BringForward;
            switch (msoZOrderCmd)
            {
                case MsoZOrderCmd.msoBringInFrontOfText:
                    zOrderCommand = ZOrderCommand.BringInFrontOfText;
                    break;
                case MsoZOrderCmd.msoBringToFront:
                    zOrderCommand = ZOrderCommand.BringToFront;
                    break;
                case MsoZOrderCmd.msoSendBackward:
                    zOrderCommand = ZOrderCommand.SendBackward;
                    break;
                case MsoZOrderCmd.msoSendBehindText:
                    zOrderCommand = ZOrderCommand.SendBehindText;
                    break;
                case MsoZOrderCmd.msoSendToBack:
                    zOrderCommand = ZOrderCommand.SendToBack;
                    break;
                default:
                    zOrderCommand = ZOrderCommand.BringForward;
                    break;
            }
            return zOrderCommand;
        }

        public static MsoZOrderCmd ConvertBack(ZOrderCommand zOrderCommand)
        {
            MsoZOrderCmd msoZOrderCmd = MsoZOrderCmd.msoBringForward;
            switch (zOrderCommand)
            {
                case ZOrderCommand.BringInFrontOfText:
                    msoZOrderCmd = MsoZOrderCmd.msoBringInFrontOfText;
                    break;
                case ZOrderCommand.BringToFront:
                    msoZOrderCmd = MsoZOrderCmd.msoBringToFront;
                    break;
                case ZOrderCommand.SendBackward:
                    msoZOrderCmd = MsoZOrderCmd.msoSendBackward;
                    break;
                case ZOrderCommand.SendBehindText:
                    msoZOrderCmd = MsoZOrderCmd.msoSendBehindText;
                    break;
                case ZOrderCommand.SendToBack:
                    msoZOrderCmd = MsoZOrderCmd.msoSendToBack;
                    break;
                default:
                    msoZOrderCmd = MsoZOrderCmd.msoBringForward;
                    break;
            }
            return msoZOrderCmd;
        }
    }
}
