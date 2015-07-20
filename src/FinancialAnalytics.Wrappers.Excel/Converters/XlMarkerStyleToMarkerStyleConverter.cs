using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
	public class XlMarkerStyleToMarkerStyleConverter
	{
		public static MarkerStyle Convert(XlMarkerStyle xlMarkerStyle)
		{
			MarkerStyle markerStyle;
			switch (xlMarkerStyle)
			{
				case XlMarkerStyle.xlMarkerStyleAutomatic:
					markerStyle = MarkerStyle.MarkerStyleAutomatic;
					break;
				case XlMarkerStyle.xlMarkerStyleCircle:
					markerStyle = MarkerStyle.MarkerStyleCircle;
					break;
				case XlMarkerStyle.xlMarkerStyleDash:
					markerStyle = MarkerStyle.MarkerStyleDash;
					break;
				case XlMarkerStyle.xlMarkerStyleDiamond:
					markerStyle = MarkerStyle.MarkerStyleDiamond;
					break;
				case XlMarkerStyle.xlMarkerStyleDot:
					markerStyle = MarkerStyle.MarkerStyleDot;
					break;
				case XlMarkerStyle.xlMarkerStyleNone:
					markerStyle = MarkerStyle.MarkerStyleNone;
					break;
				case XlMarkerStyle.xlMarkerStylePicture:
					markerStyle = MarkerStyle.MarkerStylePicture;
					break;
				case XlMarkerStyle.xlMarkerStylePlus:
					markerStyle = MarkerStyle.MarkerStylePlus;
					break;
				case XlMarkerStyle.xlMarkerStyleSquare:
					markerStyle = MarkerStyle.MarkerStyleSquare;
					break;
				case XlMarkerStyle.xlMarkerStyleStar:
					markerStyle = MarkerStyle.MarkerStyleStar;
					break;
				case XlMarkerStyle.xlMarkerStyleTriangle:
					markerStyle = MarkerStyle.MarkerStyleTriangle;
					break;
				case XlMarkerStyle.xlMarkerStyleX:
					markerStyle = MarkerStyle.MarkerStyleX;
					break;
				default:
					markerStyle = MarkerStyle.MarkerStyleNone;
					break;
			}
			return markerStyle;
		}

		public static XlMarkerStyle Convert(MarkerStyle markerStyle)
		{
			XlMarkerStyle xlMarkerStyle;
			switch (markerStyle)
			{
				case MarkerStyle.MarkerStyleAutomatic:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStyleAutomatic;
					break;
				case MarkerStyle.MarkerStyleCircle:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
					break;
				case MarkerStyle.MarkerStyleDash:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStyleDash;
					break;
				case MarkerStyle.MarkerStyleDiamond:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStyleDiamond;
					break;
				case MarkerStyle.MarkerStyleDot:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStyleDot;
					break;
				case MarkerStyle.MarkerStyleNone:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
					break;
				case MarkerStyle.MarkerStylePicture:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStylePicture;
					break;
				case MarkerStyle.MarkerStylePlus:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStylePlus;
					break;
				case MarkerStyle.MarkerStyleSquare:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStyleSquare;
					break;
				case MarkerStyle.MarkerStyleStar:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStyleStar;
					break;
				case MarkerStyle.MarkerStyleTriangle:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStyleTriangle;
					break;
				case MarkerStyle.MarkerStyleX:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStyleX;
					break;
				default:
					xlMarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
					break;
			}
			return xlMarkerStyle;
		}
	}
}	
