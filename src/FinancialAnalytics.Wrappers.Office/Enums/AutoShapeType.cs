﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Enums
{
    public enum AutoShapeType
    {
		ShapeMixed = -2,
		ShapeRectangle = 1,
		ShapeParallelogram = 2,
		ShapeTrapezoid = 3,
		ShapeDiamond = 4,
		ShapeRoundedRectangle = 5,
		ShapeOctagon = 6,
		ShapeIsoscelesTriangle = 7,
		ShapeRightTriangle = 8,
		ShapeOval = 9,
		ShapeHexagon = 10,
		ShapeCross = 11,
		ShapeRegularPentagon = 12,
		ShapeCan = 13,
		ShapeCube = 14,
		ShapeBevel = 15,
		ShapeFoldedCorner = 16,
		ShapeSmileyFace = 17,
		ShapeDonut = 18,
		ShapeNoSymbol = 19,
		ShapeBlockArc = 20,
		ShapeHeart = 21,
		ShapeLightningBolt = 22,
		ShapeSun = 23,
		ShapeMoon = 24,
		ShapeArc = 25,
		ShapeDoubleBracket = 26,
		ShapeDoubleBrace = 27,
		ShapePlaque = 28,
		ShapeLeftBracket = 29,
		ShapeRightBracket = 30,
		ShapeLeftBrace = 31,
		ShapeRightBrace = 32,
		ShapeRightArrow = 33,
		ShapeLeftArrow = 34,
		ShapeUpArrow = 35,
		ShapeDownArrow = 36,
		ShapeLeftRightArrow = 37,
		ShapeUpDownArrow = 38,
		ShapeQuadArrow = 39,
		ShapeLeftRightUpArrow = 40,
		ShapeBentArrow = 41,
		ShapeUTurnArrow = 42,
		ShapeLeftUpArrow = 43,
		ShapeBentUpArrow = 44,
		ShapeCurvedRightArrow = 45,
		ShapeCurvedLeftArrow = 46,
		ShapeCurvedUpArrow = 47,
		ShapeCurvedDownArrow = 48,
		ShapeStripedRightArrow = 49,
		ShapeNotchedRightArrow = 50,
		ShapePentagon = 51,
		ShapeChevron = 52,
		ShapeRightArrowCallout = 53,
		ShapeLeftArrowCallout = 54,
		ShapeUpArrowCallout = 55,
		ShapeDownArrowCallout = 56,
		ShapeLeftRightArrowCallout = 57,
		ShapeUpDownArrowCallout = 58,
		ShapeQuadArrowCallout = 59,
		ShapeCircularArrow = 60,
		ShapeFlowchartProcess = 61,
		ShapeFlowchartAlternateProcess = 62,
		ShapeFlowchartDecision = 63,
		ShapeFlowchartData = 64,
		ShapeFlowchartPredefinedProcess = 65,
		ShapeFlowchartInternalStorage = 66,
		ShapeFlowchartDocument = 67,
		ShapeFlowchartMultidocument = 68,
		ShapeFlowchartTerminator = 69,
		ShapeFlowchartPreparation = 70,
		ShapeFlowchartManualInput = 71,
		ShapeFlowchartManualOperation = 72,
		ShapeFlowchartConnector = 73,
		ShapeFlowchartOffpageConnector = 74,
		ShapeFlowchartCard = 75,
		ShapeFlowchartPunchedTape = 76,
		ShapeFlowchartSummingJunction = 77,
		ShapeFlowchartOr = 78,
		ShapeFlowchartCollate = 79,
		ShapeFlowchartSort = 80,
		ShapeFlowchartExtract = 81,
		ShapeFlowchartMerge = 82,
		ShapeFlowchartStoredData = 83,
		ShapeFlowchartDelay = 84,
		ShapeFlowchartSequentialAccessStorage = 85,
		ShapeFlowchartMagneticDisk = 86,
		ShapeFlowchartDirectAccessStorage = 87,
		ShapeFlowchartDisplay = 88,
		ShapeExplosion1 = 89,
		ShapeExplosion2 = 90,
		Shape4pointStar = 91,
		Shape5pointStar = 92,
		Shape8pointStar = 93,
		Shape16pointStar = 94,
		Shape24pointStar = 95,
		Shape32pointStar = 96,
		ShapeUpRibbon = 97,
		ShapeDownRibbon = 98,
		ShapeCurvedUpRibbon = 99,
		ShapeCurvedDownRibbon = 100,
		ShapeVerticalScroll = 101,
		ShapeHorizontalScroll = 102,
		ShapeWave = 103,
		ShapeDoubleWave = 104,
		ShapeRectangularCallout = 105,
		ShapeRoundedRectangularCallout = 106,
		ShapeOvalCallout = 107,
		ShapeCloudCallout = 108,
		ShapeLineCallout1 = 109,
		ShapeLineCallout2 = 110,
		ShapeLineCallout3 = 111,
		ShapeLineCallout4 = 112,
		ShapeLineCallout1AccentBar = 113,
		ShapeLineCallout2AccentBar = 114,
		ShapeLineCallout3AccentBar = 115,
		ShapeLineCallout4AccentBar = 116,
		ShapeLineCallout1NoBorder = 117,
		ShapeLineCallout2NoBorder = 118,
		ShapeLineCallout3NoBorder = 119,
		ShapeLineCallout4NoBorder = 120,
		ShapeLineCallout1BorderandAccentBar = 121,
		ShapeLineCallout2BorderandAccentBar = 122,
		ShapeLineCallout3BorderandAccentBar = 123,
		ShapeLineCallout4BorderandAccentBar = 124,
		ShapeActionButtonCustom = 125,
		ShapeActionButtonHome = 126,
		ShapeActionButtonHelp = 127,
		ShapeActionButtonInformation = 128,
		ShapeActionButtonBackorPrevious = 129,
		ShapeActionButtonForwardorNext = 130,
		ShapeActionButtonBeginning = 131,
		ShapeActionButtonEnd = 132,
		ShapeActionButtonReturn = 133,
		ShapeActionButtonDocument = 134,
		ShapeActionButtonSound = 135,
		ShapeActionButtonMovie = 136,
		ShapeBalloon = 137,
		ShapeNotPrimitive = 138,
		ShapeFlowchartOfflineStorage = 139,
		ShapeLeftRightRibbon = 140,
		ShapeDiagonalStripe = 141,
		ShapePie = 142,
		ShapeNonIsoscelesTrapezoid = 143,
		ShapeDecagon = 144,
		ShapeHeptagon = 145,
		ShapeDodecagon = 146,
		Shape6pointStar = 147,
		Shape7pointStar = 148,
		Shape10pointStar = 149,
		Shape12pointStar = 150,
		ShapeRound1Rectangle = 151,
		ShapeRound2SameRectangle = 152,
		ShapeRound2DiagRectangle = 153,
		ShapeSnipRoundRectangle = 154,
		ShapeSnip1Rectangle = 155,
		ShapeSnip2SameRectangle = 156,
		ShapeSnip2DiagRectangle = 157,
		ShapeFrame = 158,
		ShapeHalfFrame = 159,
		ShapeTear = 160,
		ShapeChord = 161,
		ShapeCorner = 162,
		ShapeMathPlus = 163,
		ShapeMathMinus = 164,
		ShapeMathMultiply = 165,
		ShapeMathDivide = 166,
		ShapeMathEqual = 167,
		ShapeMathNotEqual = 168,
		ShapeCornerTabs = 169,
		ShapeSquareTabs = 170,
		ShapePlaqueTabs = 171,
		ShapeGear6 = 172,
		ShapeGear9 = 173,
		ShapeFunnel = 174,
		ShapePieWedge = 175,
		ShapeLeftCircularArrow = 176,
		ShapeLeftRightCircularArrow = 177,
		ShapeSwooshArrow = 178,
		ShapeCloud = 179,
		ShapeChartX = 180,
		ShapeChartStar = 181,
		ShapeChartPlus = 182,
		ShapeLineInverse = 183
    }
}
