using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Rg = Rhino.Geometry;

namespace Bumblebee
{
    public static class GhExtensions
    {

        public static bool TryGetApp(this IGH_Goo input, ref ExApp application)
        {
            ExApp app;
            ExWorkbook book;
            ExWorksheet sheet;
            ExRange range;

            if (input == null)
            {
                application = new ExApp();
                return true;
            }
            else if (input.CastTo<ExWorkbook>(out book))
            {
                application = new ExApp(book.ParentApp);
                return true;
            }
            else if (input.CastTo<ExWorksheet>(out sheet))
            {
                application = new ExApp(sheet.ParentApp);
                return true;
            }
            else if (input.CastTo<ExRange>(out range))
            {
                application = range.ParentApp;
                return true;
            }
            else if (input.CastTo<ExApp>(out app))
            {
                application = new ExApp(app);
                return true;
            }
            else
            {
                application = new ExApp();
                return true;
            }
        }

        public static bool TryGetWorkbook(this IGH_Goo input, ref ExWorkbook workbook, ExApp application = null)
        {
            ExApp app;
            ExWorkbook book;
            ExWorksheet sheet;
            ExRange range;
            string name;

            if (input == null)
            {
                if (application == null)
                {
                    app = new ExApp();
                }
                else
                {
                    app = new ExApp(application);
                }
                workbook = app.GetActiveWorkbook();
                return true;
            }
            else if (input.CastTo<ExWorkbook>(out book))
            {
                workbook = new ExWorkbook(book);
                return true;
            }
            else if (input.CastTo<ExWorksheet>(out sheet))
            {
                workbook = sheet.Workbook;
                return true;
            }
            else if (input.CastTo<ExRange>(out range))
            {
                workbook = range.Workbook;
                return true;
            }
            else if (input.CastTo<ExApp>(out app))
            {
                workbook = app.GetActiveWorkbook();
                return true;
            }
            else if (input.CastTo<string>(out name))
            {
                if (application == null)
                {
                    app = new ExApp();
                }
                else
                {
                    app = new ExApp(application);
                }
                workbook = app.GetWorkbook(name);
                return true;
            }
            else
            {
                app = new ExApp();
                workbook = app.GetActiveWorkbook();
                return true;
            }
        }

        public static bool TryGetWorksheet(this IGH_Goo input, ref ExWorksheet worksheet, ExWorkbook book = null)
        {
            ExApp app;
            ExWorkbook workbook;
            ExWorksheet sheet;
            ExRange range;
            string name;

            if (input == null)
            {
                if (book == null)
                {
                    worksheet = ExWorksheet.ActiveWorksheet();
                    return true;
                }
                else
                {
                    worksheet = book.GetActiveWorksheet();
                    return true;
                }
            }
            else if (input.CastTo<ExWorksheet>(out sheet))
            {
                worksheet = new ExWorksheet(sheet);
                return true;
            }
            else if (input.CastTo<ExWorkbook>(out workbook))
            {
                worksheet = workbook.GetActiveWorksheet();
                return true;
            }
            else if (input.CastTo<ExApp>(out app))
            {
                worksheet = app.GetActiveWorksheet();
                return true;
            }
            else if (input.CastTo<ExRange>(out range))
            {
                worksheet = range.Worksheet;
                return true;
            }
            else if (input.CastTo<string>(out name))
            {
                if (book == null)
                {
                    worksheet = ExWorksheet.ActiveWorksheet(name);
                    return true;
                }
                else
                {
                    worksheet = book.GetWorksheet(name);
                    return true;
                }
            }
            else
            {
                    worksheet = ExWorksheet.ActiveWorksheet();
                    return true;
            }
        }

        public static bool TryGetRange(this IGH_Goo input, ref ExRange range, ExWorksheet worksheet = null)
        {
            ExApp app;
            ExWorkbook book;
            ExWorksheet sheet;
            ExRange rng;
            ExCell cell;
            string address;

            if(input == null)
            {
                if (worksheet.ComObj == null)
                {
                    range = GetActiveRange();
                }
                else
                {
                    range = worksheet.GetUsedRange();
                }
                return true;
            }
            else if (input.CastTo<ExApp>(out app))
            {
                sheet = app.GetActiveWorksheet();
                range = sheet.GetUsedRange();
                return true;
            }
            else if (input.CastTo<ExWorkbook>(out book))
            {
                sheet = book.GetActiveWorksheet();
                range = sheet.GetUsedRange();
                return true;
            }
            else if (input.CastTo<ExWorksheet>(out sheet))
            {
                range = sheet.GetUsedRange();
                return true;
            }
            else if (input.CastTo<ExRange>(out rng))
            {
                range = new ExRange(rng);
                return true;
            }
            else if (input.CastTo<ExCell>(out cell))
            {
                if (worksheet.ComObj == null)
                {
                    range = GetActiveRange();
                }
                else
                {
                    range = worksheet.GetRange(cell, cell);
                }
                return true;
            }
            else if (input.CastTo<string>(out address))
            {
                if (worksheet.ComObj == null)
                {
                    range = GetActiveRange();
                }
                else
                {
                    range = worksheet.GetRange(address);
                }
                return true;
            }
            else
            {
                range = GetActiveRange();
                return true;
            }
        }

        public static ExRange GetActiveRange()
        {
                ExApp app = new ExApp();
                ExWorksheet sheet = app.GetActiveWorksheet();
                return sheet.GetUsedRange();
        }

        public static bool TryGetCell(this IGH_Goo input, ref ExCell cell)
        {
            ExCell cl;
            GH_Point gpt;
            Rg.Point3d p3d;
            Rg.Point2d p2d;
            string address;
            Rg.Interval domain;

            if (input == null) return false;

            if (input.CastTo<ExCell>(out cl))
            {
                cell = new ExCell(cl);
                return true;
            }
            else if (input.CastTo<GH_Point>(out gpt))
            {
                cell = new ExCell((int)gpt.Value.X, (int)gpt.Value.Y);
                return true;
            }
            else if (input.CastTo<Rg.Point3d>(out p3d))
            {
                cell = new ExCell((int)p3d.X, (int)p3d.Y);
                return true;
            }
            else if (input.CastTo<Rg.Point2d>(out p2d))
            {
                cell = new ExCell((int)p2d.X, (int)p2d.Y);
                return true;
            }
            else if (input.CastTo<Rg.Interval>(out domain))
            {
                cell = new ExCell((int)domain.T0, (int)domain.T1);
                return true;
            }
            else if (input.CastTo<string>(out address))
            {
                cell = new ExCell(address);
                return true;
            }
            else
            {
                cell = null;
                return false;
            }
        }

    }
}
