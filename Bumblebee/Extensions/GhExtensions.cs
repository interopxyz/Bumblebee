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

        public static ExWorkbook ToWorkbook(this IGH_Goo input)
        {
            ExApp app;
            ExWorkbook workbook;

            if (input.CastTo<ExWorkbook>(out workbook))
            {
                return workbook;
            }
            else if (input.CastTo<ExApp>(out app))
            {
                return app.GetActiveWorkbook();
            }
            else
            {
                return null;
            }
        }

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

        public static bool TryGetRange(this IGH_Goo input, ref ExRange range)
        {
            ExRange rng;
            ExCell cell;
            string address;

            if (input.CastTo<ExRange>(out rng))
            {
                range = new ExRange(rng);
                return true;
            }
            else if (input.CastTo<ExCell>(out cell))
            {
                range = new ExRange(cell,cell);
                return true;
            }
            else if (input.CastTo<string>(out address))
            {
                range = new ExRange(address);
                return true;
            }
            else
            {
                range = null;
                return false;
            }
        }

    }
}
