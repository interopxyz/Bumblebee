using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = Microsoft.Office.Interop.Excel;
using Vi = Microsoft.Vbe.Interop;
using Mc = Microsoft.Office.Core;

namespace Bumblebee
{
    public static class ExExtensions
    {
        public static XL.XlFileFormat ToExcel(this Extensions input)
        {
            switch (input)
            {
                default:
                    return XL.XlFileFormat.xlWorkbookDefault;
                case Extensions.xlxm:
                    return XL.XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
            }
        }
        public static LineType ToBB(this XL.XlLineStyle input)
        {
            switch (input)
            {
                case XL.XlLineStyle.xlContinuous:
                    return LineType.Continuous;
                case XL.XlLineStyle.xlDash:
                    return LineType.Dash;
                case XL.XlLineStyle.xlDashDot:
                    return LineType.DashDot;
                case XL.XlLineStyle.xlDashDotDot:
                    return LineType.DashDotDot;
                case XL.XlLineStyle.xlDot:
                    return LineType.Dot;
                case XL.XlLineStyle.xlDouble:
                    return LineType.Double;
                case XL.XlLineStyle.xlSlantDashDot:
                    return LineType.SlantDashDot;
                default:
                    return LineType.None;
            }
        }

        public static Mc.MsoArrowheadStyle ToExcel(this ArrowStyle input)
        {
            switch(input)
            {
                default:
                    return Mc.MsoArrowheadStyle.msoArrowheadNone;
                case ArrowStyle.Diamond:
                    return Mc.MsoArrowheadStyle.msoArrowheadDiamond;
                case ArrowStyle.Open:
                    return Mc.MsoArrowheadStyle.msoArrowheadOpen;
                case ArrowStyle.Oval:
                    return Mc.MsoArrowheadStyle.msoArrowheadOval;
                case ArrowStyle.Stealth:
                    return Mc.MsoArrowheadStyle.msoArrowheadStealth;
                case ArrowStyle.Triangle:
                    return Mc.MsoArrowheadStyle.msoArrowheadTriangle;
            }
        }

        public static Vi.vbext_ComponentType ToExcel(this VbModuleType input)
        {
            switch (input)
            {
                case VbModuleType.ClassModule:
                    return Vi.vbext_ComponentType.vbext_ct_ClassModule;
                case VbModuleType.Document:
                    return Vi.vbext_ComponentType.vbext_ct_Document;
                case VbModuleType.MSForm:
                    return Vi.vbext_ComponentType.vbext_ct_MSForm;
                case VbModuleType.ActiveX:
                    return Vi.vbext_ComponentType.vbext_ct_ActiveXDesigner;
                default:
                    return Vi.vbext_ComponentType.vbext_ct_StdModule;
            }
        }

        public static XL.XlLineStyle ToExcel(this LineType input)
        {
            switch (input)
            {
                case LineType.Continuous:
                    return XL.XlLineStyle.xlContinuous;
                case LineType.Dash:
                    return XL.XlLineStyle.xlDash;
                case LineType.DashDot:
                    return XL.XlLineStyle.xlDashDot;
                case LineType.DashDotDot:
                    return XL.XlLineStyle.xlDashDotDot;
                case LineType.Dot:
                    return XL.XlLineStyle.xlDot;
                case LineType.Double:
                    return XL.XlLineStyle.xlDouble;
                case LineType.SlantDashDot:
                    return XL.XlLineStyle.xlSlantDashDot;
                default:
                    return XL.XlLineStyle.xlLineStyleNone;
            }
        }

        public static XL.XlHAlign ToExcelHalign(this Justification input)
        {
            switch (input)
            {
                case Justification.BottomMiddle:
                case Justification.CenterMiddle:
                case Justification.TopMiddle:
                    return XL.XlHAlign.xlHAlignCenter;
                case Justification.BottomRight:
                case Justification.CenterRight:
                case Justification.TopRight:
                    return XL.XlHAlign.xlHAlignRight;
                default:
                    return XL.XlHAlign.xlHAlignLeft;
            }
        }

        public static XL.XlVAlign ToExcelValign(this Justification input)
        {
            switch (input)
            {
                case Justification.CenterLeft:
                case Justification.CenterMiddle:
                case Justification.CenterRight:
                    return XL.XlVAlign.xlVAlignCenter;
                case Justification.TopLeft:
                case Justification.TopMiddle:
                case Justification.TopRight:
                    return XL.XlVAlign.xlVAlignTop;
                default:
                    return XL.XlVAlign.xlVAlignBottom;
            }
        }

        public static XL.XlBorderWeight ToExcel(this BorderWeight input)
        {
            switch (input)
            {
                case BorderWeight.Hairline:
                    return XL.XlBorderWeight.xlHairline;
                case BorderWeight.Medium:
                    return XL.XlBorderWeight.xlMedium;
                case BorderWeight.Thick:
                    return XL.XlBorderWeight.xlThick;
                default:
                    return XL.XlBorderWeight.xlThin;
            }
        }

        public static BorderWeight ToBB(this XL.XlBorderWeight input)
        {
            switch (input)
            {
                case XL.XlBorderWeight.xlHairline:
                    return BorderWeight.Hairline;
                case XL.XlBorderWeight.xlMedium:
                    return BorderWeight.Medium;
                case XL.XlBorderWeight.xlThick:
                    return BorderWeight.Thick;
                default:
                    return BorderWeight.Thin;
            }
        }

        public static Mc.MsoAutoShapeType ToExcel(this ShapeArrow input)
        {
            switch (input)
            {
                case ShapeArrow.Circular:
                    return Mc.MsoAutoShapeType.msoShapeCircularArrow;
                case ShapeArrow.Left:
                    return Mc.MsoAutoShapeType.msoShapeLeftArrow;
                case ShapeArrow.Up:
                    return Mc.MsoAutoShapeType.msoShapeUpArrow;
                case ShapeArrow.Down:
                    return Mc.MsoAutoShapeType.msoShapeDownArrow;
                case ShapeArrow.LeftRight:
                    return Mc.MsoAutoShapeType.msoShapeLeftRightArrow;
                case ShapeArrow.UpDown:
                    return Mc.MsoAutoShapeType.msoShapeUpDownArrow;
                case ShapeArrow.Quad:
                    return Mc.MsoAutoShapeType.msoShapeQuadArrow;
                case ShapeArrow.LeftRightUp:
                    return Mc.MsoAutoShapeType.msoShapeLeftRightUpArrow;
                case ShapeArrow.Bent:
                    return Mc.MsoAutoShapeType.msoShapeBentArrow;
                case ShapeArrow.UTurn:
                    return Mc.MsoAutoShapeType.msoShapeUTurnArrow;
                case ShapeArrow.LeftUp:
                    return Mc.MsoAutoShapeType.msoShapeLeftUpArrow;
                case ShapeArrow.BentUp:
                    return Mc.MsoAutoShapeType.msoShapeBentUpArrow;
                case ShapeArrow.CurvedRight:
                    return Mc.MsoAutoShapeType.msoShapeCurvedRightArrow;
                case ShapeArrow.CurvedLeft:
                    return Mc.MsoAutoShapeType.msoShapeCurvedLeftArrow;
                case ShapeArrow.CurvedUp:
                    return Mc.MsoAutoShapeType.msoShapeCurvedUpArrow;
                case ShapeArrow.CurvedDown:
                    return Mc.MsoAutoShapeType.msoShapeCurvedDownArrow;
                case ShapeArrow.StripedRight:
                    return Mc.MsoAutoShapeType.msoShapeStripedRightArrow;
                case ShapeArrow.NotchedRight:
                    return Mc.MsoAutoShapeType.msoShapeNotchedRightArrow;
                case ShapeArrow.Swoosh:
                    return Mc.MsoAutoShapeType.msoShapeSwooshArrow;
                case ShapeArrow.LeftCircular:
                    return Mc.MsoAutoShapeType.msoShapeLeftCircularArrow;
                case ShapeArrow.LeftRightCircular:
                    return Mc.MsoAutoShapeType.msoShapeLeftRightCircularArrow;
                default:
                    return Mc.MsoAutoShapeType.msoShapeRightArrow;
            }
        }

        public static Mc.MsoAutoShapeType ToExcel(this ShapeStar input)
        {
            switch (input)
            {
                case ShapeStar.Pt5:
                    return Mc.MsoAutoShapeType.msoShape5pointStar;
                case ShapeStar.Pt6:
                    return Mc.MsoAutoShapeType.msoShape6pointStar;
                case ShapeStar.Pt7:
                    return Mc.MsoAutoShapeType.msoShape7pointStar;
                case ShapeStar.Pt8:
                    return Mc.MsoAutoShapeType.msoShape8pointStar;
                case ShapeStar.Pt10:
                    return Mc.MsoAutoShapeType.msoShape10pointStar;
                case ShapeStar.Pt12:
                    return Mc.MsoAutoShapeType.msoShape12pointStar;
                case ShapeStar.Pt16:
                    return Mc.MsoAutoShapeType.msoShape16pointStar;
                case ShapeStar.Pt24:
                    return Mc.MsoAutoShapeType.msoShape24pointStar;
                case ShapeStar.Pt32:
                    return Mc.MsoAutoShapeType.msoShape32pointStar;
                default:
                    return Mc.MsoAutoShapeType.msoShape4pointStar;
            }
        }

        public static Mc.MsoAutoShapeType ToExcel(this ShapeFlowChart input)
        {
            switch (input)
            {
                case ShapeFlowChart.AlternateProcess:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartAlternateProcess;
                case ShapeFlowChart.Collate:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartCollate;
                case ShapeFlowChart.Connector:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartConnector;
                case ShapeFlowChart.Data:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartData;
                case ShapeFlowChart.Decision:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartDecision;
                case ShapeFlowChart.Delay:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartDelay;
                case ShapeFlowChart.DirectAccessStorage:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartDirectAccessStorage;
                case ShapeFlowChart.Display:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartDisplay;
                case ShapeFlowChart.Document:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartDocument;
                case ShapeFlowChart.Extract:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartExtract;
                case ShapeFlowChart.InternalStorage:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartInternalStorage;
                case ShapeFlowChart.MagneticDisk:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartMagneticDisk;
                case ShapeFlowChart.ManualInput:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartManualInput;
                case ShapeFlowChart.ManualOperation:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartManualOperation;
                case ShapeFlowChart.Merge:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartMerge;
                case ShapeFlowChart.Multidocument:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartMultidocument;
                case ShapeFlowChart.OffpageConnector:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartOffpageConnector;
                case ShapeFlowChart.OfflineStorage:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartOfflineStorage;
                case ShapeFlowChart.Or:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartOr;
                case ShapeFlowChart.PredefinedProcess:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartPredefinedProcess;
                case ShapeFlowChart.Preparation:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartPreparation;
                case ShapeFlowChart.Process:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartProcess;
                case ShapeFlowChart.PunchedTape:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartPunchedTape;
                case ShapeFlowChart.SequentialAccessStorage:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartSequentialAccessStorage;
                case ShapeFlowChart.Sort:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartSort;
                case ShapeFlowChart.StoredData:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartStoredData;
                case ShapeFlowChart.SummingJunction:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartSummingJunction;
                case ShapeFlowChart.Terminator:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartTerminator;
                default:
                    return Mc.MsoAutoShapeType.msoShapeFlowchartCard;
            }
        }

        public static Mc.MsoAutoShapeType ToExcel(this ShapeSymbol input)
        {
            switch (input)
            {
                case ShapeSymbol.Plus:
                    return Mc.MsoAutoShapeType.msoShapeMathPlus;
                case ShapeSymbol.Minus:
                    return Mc.MsoAutoShapeType.msoShapeMathMinus;
                case ShapeSymbol.Multiply:
                    return Mc.MsoAutoShapeType.msoShapeMathMultiply;
                case ShapeSymbol.Divide:
                    return Mc.MsoAutoShapeType.msoShapeMathDivide;
                case ShapeSymbol.NotEqual:
                    return Mc.MsoAutoShapeType.msoShapeMathNotEqual;
                case ShapeSymbol.LeftBracket:
                    return Mc.MsoAutoShapeType.msoShapeLeftBracket;
                case ShapeSymbol.RightBracket:
                    return Mc.MsoAutoShapeType.msoShapeRightBracket;
                case ShapeSymbol.DoubleBracket:
                    return Mc.MsoAutoShapeType.msoShapeDoubleBracket;
                case ShapeSymbol.LeftBrace:
                    return Mc.MsoAutoShapeType.msoShapeLeftBrace;
                case ShapeSymbol.RightBrace:
                    return Mc.MsoAutoShapeType.msoShapeRightBrace;
                case ShapeSymbol.DoubleBrace:
                    return Mc.MsoAutoShapeType.msoShapeDoubleBrace;
                default:
                    return Mc.MsoAutoShapeType.msoShapeMathEqual;
            }
        }

        public static Mc.MsoAutoShapeType ToExcel(this ShapeGeometry input)
        {
            switch (input)
            {
                case ShapeGeometry.BlockArc:
                    return Mc.MsoAutoShapeType.msoShapeBlockArc;
                case ShapeGeometry.Cross:
                    return Mc.MsoAutoShapeType.msoShapeCross;
                case ShapeGeometry.Decagon:
                    return Mc.MsoAutoShapeType.msoShapeDecagon;
                case ShapeGeometry.Diamond:
                    return Mc.MsoAutoShapeType.msoShapeDiamond;
                case ShapeGeometry.Dodecagon:
                    return Mc.MsoAutoShapeType.msoShapeDodecagon;
                case ShapeGeometry.Donut:
                    return Mc.MsoAutoShapeType.msoShapeDonut;
                case ShapeGeometry.Heptagon:
                    return Mc.MsoAutoShapeType.msoShapeHeptagon;
                case ShapeGeometry.Hexagon:
                    return Mc.MsoAutoShapeType.msoShapeHexagon;
                case ShapeGeometry.IsoscelesTriangle:
                    return Mc.MsoAutoShapeType.msoShapeIsoscelesTriangle;
                case ShapeGeometry.NonIsoscelesTrapezoid:
                    return Mc.MsoAutoShapeType.msoShapeNonIsoscelesTrapezoid;
                case ShapeGeometry.Octagon:
                    return Mc.MsoAutoShapeType.msoShapeOctagon;
                case ShapeGeometry.Oval:
                    return Mc.MsoAutoShapeType.msoShapeOval;
                case ShapeGeometry.Parallelogram:
                    return Mc.MsoAutoShapeType.msoShapeParallelogram;
                case ShapeGeometry.Pentagon:
                    return Mc.MsoAutoShapeType.msoShapePentagon;
                case ShapeGeometry.RegularPentagon:
                    return Mc.MsoAutoShapeType.msoShapeRegularPentagon;
                case ShapeGeometry.RightTriangle:
                    return Mc.MsoAutoShapeType.msoShapeRightTriangle;
                case ShapeGeometry.Round1Rectangle:
                    return Mc.MsoAutoShapeType.msoShapeRound1Rectangle;
                case ShapeGeometry.Round2DiagRectangle:
                    return Mc.MsoAutoShapeType.msoShapeRound2DiagRectangle;
                case ShapeGeometry.Round2SameRectangle:
                    return Mc.MsoAutoShapeType.msoShapeRound2SameRectangle;
                case ShapeGeometry.RoundedRectangle:
                    return Mc.MsoAutoShapeType.msoShapeRoundedRectangle;
                case ShapeGeometry.Snip1Rectangle:
                    return Mc.MsoAutoShapeType.msoShapeSnip1Rectangle;
                case ShapeGeometry.Snip2DiagRectangle:
                    return Mc.MsoAutoShapeType.msoShapeSnip2DiagRectangle;
                case ShapeGeometry.Snip2SameRectangle:
                    return Mc.MsoAutoShapeType.msoShapeSnip2SameRectangle;
                case ShapeGeometry.SnipRoundRectangle:
                    return Mc.MsoAutoShapeType.msoShapeSnipRoundRectangle;
                case ShapeGeometry.Trapezoid:
                    return Mc.MsoAutoShapeType.msoShapeTrapezoid;
                default:
                    return Mc.MsoAutoShapeType.msoShapeRectangle;
            }
        }

        public static Mc.MsoAutoShapeType ToExcel(this ShapeFigure input)
        {
            switch (input)
            {
                case ShapeFigure.Arc:
                    return Mc.MsoAutoShapeType.msoShapeArc;
                case ShapeFigure.Balloon:
                    return Mc.MsoAutoShapeType.msoShapeBalloon;
                case ShapeFigure.Bevel:
                    return Mc.MsoAutoShapeType.msoShapeBevel;
                case ShapeFigure.Can:
                    return Mc.MsoAutoShapeType.msoShapeCan;
                case ShapeFigure.Chevron:
                    return Mc.MsoAutoShapeType.msoShapeChevron;
                case ShapeFigure.Chord:
                    return Mc.MsoAutoShapeType.msoShapeChord;
                case ShapeFigure.Cloud:
                    return Mc.MsoAutoShapeType.msoShapeCloud;
                case ShapeFigure.Corner:
                    return Mc.MsoAutoShapeType.msoShapeCorner;
                case ShapeFigure.DiagonalStripe:
                    return Mc.MsoAutoShapeType.msoShapeDiagonalStripe;
                case ShapeFigure.DoubleWave:
                    return Mc.MsoAutoShapeType.msoShapeDoubleWave;
                case ShapeFigure.Explosion1:
                    return Mc.MsoAutoShapeType.msoShapeExplosion1;
                case ShapeFigure.Explosion2:
                    return Mc.MsoAutoShapeType.msoShapeExplosion2;
                case ShapeFigure.FoldedCorner:
                    return Mc.MsoAutoShapeType.msoShapeFoldedCorner;
                case ShapeFigure.Frame:
                    return Mc.MsoAutoShapeType.msoShapeFrame;
                case ShapeFigure.Funnel:
                    return Mc.MsoAutoShapeType.msoShapeFunnel;
                case ShapeFigure.Gear6:
                    return Mc.MsoAutoShapeType.msoShapeGear6;
                case ShapeFigure.Gear9:
                    return Mc.MsoAutoShapeType.msoShapeGear9;
                case ShapeFigure.HalfFrame:
                    return Mc.MsoAutoShapeType.msoShapeHalfFrame;
                case ShapeFigure.Heart:
                    return Mc.MsoAutoShapeType.msoShapeHeart;
                case ShapeFigure.LightningBolt:
                    return Mc.MsoAutoShapeType.msoShapeLightningBolt;
                case ShapeFigure.Moon:
                    return Mc.MsoAutoShapeType.msoShapeMoon;
                case ShapeFigure.NoSymbol:
                    return Mc.MsoAutoShapeType.msoShapeNoSymbol;
                case ShapeFigure.Pie:
                    return Mc.MsoAutoShapeType.msoShapePie;
                case ShapeFigure.PieWedge:
                    return Mc.MsoAutoShapeType.msoShapePieWedge;
                case ShapeFigure.Plaque:
                    return Mc.MsoAutoShapeType.msoShapePlaque;
                case ShapeFigure.SmileyFace:
                    return Mc.MsoAutoShapeType.msoShapeSmileyFace;
                case ShapeFigure.Sun:
                    return Mc.MsoAutoShapeType.msoShapeSun;
                case ShapeFigure.Tear:
                    return Mc.MsoAutoShapeType.msoShapeTear;
                case ShapeFigure.Wave:
                    return Mc.MsoAutoShapeType.msoShapeWave;
                default:
                    return Mc.MsoAutoShapeType.msoShapeCube;
            }
        }
    }
}
