using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
	public class XlPlacementToPlacementConverter
    {
		public static PlacementType Convert(XlPlacement xlPlacement)
        {
			switch (xlPlacement)
            { 
                case XlPlacement.xlFreeFloating:
                    return PlacementType.FreeFloating;
				case XlPlacement.xlMove:
                    return PlacementType.Move;
                default:
                    return PlacementType.MoveAndSize;
            }
        }

		public static XlPlacement ConvertBack(PlacementType placement)
        {
			switch (placement)
            {
				case PlacementType.FreeFloating:
					return XlPlacement.xlFreeFloating;
				case PlacementType.Move:
					return  XlPlacement.xlMove;
                default:
                    return XlPlacement.xlMoveAndSize;
            }
        }
    }
}
