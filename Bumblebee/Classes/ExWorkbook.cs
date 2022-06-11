using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = Microsoft.Office.Interop.Excel;

namespace Bumblebee
{
    public class ExWorkbook
    {

        #region members

        public XL.Workbook ComObj = null;

        protected string name = "";

        #endregion

        #region constructors

        public ExWorkbook()
        {
        }

        public ExWorkbook(XL.Workbook comObj)
        {
            this.ComObj = comObj;
            this.name = comObj.Name;
        }

        public ExWorkbook(ExWorkbook workbook)
        {
            this.ComObj = workbook.ComObj;
            this.name = workbook.Name;
        }

        #endregion

        #region properties

        public virtual string Name
        {
            get { return name;}
        }

        public virtual ExApp ParentApp
        {
            get { return new ExApp(this.ComObj.Application); }
        }

        #endregion

        #region methods

        public void Set(XL.Workbook comObject)
        {
            this.ComObj = comObject;
            this.name = comObject.Name;
        }

        public void Save(string filename)
        {
            this.ComObj.SaveAs(filename);
            this.name = this.ComObj.Name;
        }

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

            XL.Worksheet worksheet1 = this.ComObj.Worksheets.Add();
            worksheet1.Name = name;

            return new ExWorksheet(worksheet1); ;
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

            if (this.ComObj.Worksheets.Count < 1)
            {
                output = new ExWorksheet(this.ComObj.Worksheets.Add());
            }
            else
            {
                output = new ExWorksheet(this.ComObj.ActiveSheet);
            }

            return output;
        }

        #endregion

        public void Activate()
        {
            this.ComObj.Activate();
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Workbook | "+Name;
        }

        #endregion

    }
}
