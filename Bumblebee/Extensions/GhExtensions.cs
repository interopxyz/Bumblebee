using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bumblebee
{
    public static class GhExtensions
    {

        public static ExWorksheet ToWorksheet(this IGH_Goo input)
        {
            ExApp app;
            ExWorkbook workbook;
            ExWorksheet worksheet;

            if (input.CastTo<ExWorksheet>(out worksheet))
            {
                return worksheet;
            }
            else if (input.CastTo<ExWorkbook>(out workbook))
            {
                return workbook.GetActiveWorksheet();
            }
            else if (input.CastTo<ExApp>(out app))
            {
                return app.GetActiveWorksheet();
            }
            else
            {
                return null;
            }
        }

    }
}
