using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = Microsoft.Office.Interop.Excel;

namespace Bumblebee
{
    public static class ExExtensions
    {

        public static XL.XlLineStyle ToExcel(this ExApp.LineType input)
        {
            switch (input)
            {
                case ExApp.LineType.Continuous:
                    return XL.XlLineStyle.xlContinuous;
                case ExApp.LineType.Dash:
                    return XL.XlLineStyle.xlDash;
                case ExApp.LineType.DashDot:
                    return XL.XlLineStyle.xlDashDot;
                case ExApp.LineType.DashDotDot:
                    return XL.XlLineStyle.xlDashDotDot;
                case ExApp.LineType.Dot:
                    return XL.XlLineStyle.xlDot;
                case ExApp.LineType.Double:
                    return XL.XlLineStyle.xlDouble;
                case ExApp.LineType.SlantDashDot:
                    return XL.XlLineStyle.xlSlantDashDot;
                default:
                    return XL.XlLineStyle.xlLineStyleNone;
            }
        }

        public static XL.XlHAlign ToExcelHalign(this ExApp.Justification input)
        {
            switch (input)
            {
                case ExApp.Justification.BottomMiddle:
                case ExApp.Justification.CenterMiddle:
                case ExApp.Justification.TopMiddle:
                    return XL.XlHAlign.xlHAlignCenter;
                case ExApp.Justification.BottomRight:
                case ExApp.Justification.CenterRight:
                case ExApp.Justification.TopRight:
                    return XL.XlHAlign.xlHAlignRight;
                default:
                    return XL.XlHAlign.xlHAlignLeft;
            }
        }

        public static XL.XlVAlign ToExcelValign(this ExApp.Justification input)
        {
            switch (input)
            {
                case ExApp.Justification.CenterLeft:
                case ExApp.Justification.CenterMiddle:
                case ExApp.Justification.CenterRight:
                    return XL.XlVAlign.xlVAlignCenter;
                case ExApp.Justification.TopLeft:
                case ExApp.Justification.TopMiddle:
                case ExApp.Justification.TopRight:
                    return XL.XlVAlign.xlVAlignTop;
                default:
                    return XL.XlVAlign.xlVAlignBottom;
            }
        }

        public static XL.XlBorderWeight ToExcel(this ExApp.BorderWeight input)
        {
            switch (input)
            {
                case ExApp.BorderWeight.Hairline:
                    return XL.XlBorderWeight.xlHairline;
                case ExApp.BorderWeight.Medium:
                    return XL.XlBorderWeight.xlMedium;
                case ExApp.BorderWeight.Thick:
                    return XL.XlBorderWeight.xlThick;
                default:
                    return XL.XlBorderWeight.xlThin;
            }
        }

    }
}
