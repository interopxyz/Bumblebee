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
