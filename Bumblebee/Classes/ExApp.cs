using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = Microsoft.Office.Interop.Excel;

namespace Bumblebee
{
    public class ExApp
    {

        #region members

        public XL.Application ComObj = null;

        #endregion

        #region constructors

        public ExApp()
        {
            try
            { 
                this.ComObj = (XL.Application)Marshal2.GetActiveObject("Excel.Application"); 
            }
            catch (Exception e)
            {
                this.ComObj = new XL.Application();
            }
            if (this.ComObj.ActiveWorkbook == null) this.ComObj.Workbooks.Add();
            if(! this.ComObj.Visible) this.ComObj.Visible = true;
        }

        public ExApp(ExApp exApp)
        {
            this.ComObj = exApp.ComObj;
        }

        public ExApp(XL.Application comObj)
        {
            this.ComObj = comObj;
        }

        #endregion

        #region properties



        #endregion

        #region methods

        public void Freeze()
        {
            this.ComObj.ScreenUpdating = false;
        }
        public void UnFreeze()
        {
            this.ComObj.ScreenUpdating = true;
        }

        #region workbooks

        public ExWorkbook LoadWorkbook(string filePath)
        {
            ExWorkbook workbook = new ExWorkbook( this.ComObj.Workbooks.Open(filePath));

            return workbook;
        }

        public List<ExWorkbook> GetWorkbooks()
        {
            List<ExWorkbook> output = new List<ExWorkbook>();

            foreach (XL.Workbook workbook in this.ComObj.Workbooks)
            {
                    output.Add(new ExWorkbook(workbook));
            }

            return output;
        }

        public ExWorkbook GetWorkbook(string name)
        {
            ExWorkbook output = new ExWorkbook();

            foreach (XL.Workbook workbook in this.ComObj.Workbooks)
            {
            string file = Path.ChangeExtension(workbook.Name, null);
                
                if (file == name)
                {
                    output.Set(workbook);
                    return output;
                }
            }

            return null;
        }

        public ExWorkbook GetActiveWorkbook()
        {
            ExWorkbook output = new ExWorkbook();

            if (this.ComObj.Workbooks.Count < 1)
            {
                output.Set(this.ComObj.Workbooks.Add());
            }
            else
            {
                output.Set(this.ComObj.ActiveWorkbook);
            }

            return output;
        }

        #endregion

        #region worksheets

        public ExWorksheet GetWorksheet(string name)
        {
            ExWorksheet output = new ExWorksheet();

            foreach (XL.Worksheet worksheet in this.ComObj.Worksheets)
            {
                if (worksheet.Name == name)
                {
                    output = new ExWorksheet(worksheet);
                    return output;
                }
            }

            return null;
        }

        public List<ExWorksheet> GetWorksheets()
        {
            List<ExWorksheet> worksheets = new List<ExWorksheet>();

            foreach (XL.Worksheet sheet in this.ComObj.Worksheets)
            {
                worksheets.Add(new ExWorksheet(sheet));
            }

            return worksheets;
        }

        public ExWorksheet GetActiveWorksheet()
        {
            ExWorksheet output = new ExWorksheet();

            if (this.ComObj.ActiveSheet == null)
            {
                XL.Workbook book = this.ComObj.Workbooks.Add();
                XL.Worksheet sheet = book.Worksheets.Add();
                output = new ExWorksheet(sheet);
            }
            else
            {
                output = new ExWorksheet(this.ComObj.ActiveSheet);
            }

            return output;
        }

        #endregion

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Excel | App";
        }

        #endregion

    }
}
