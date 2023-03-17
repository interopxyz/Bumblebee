using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = Microsoft.Office.Interop.Excel;
using Vi = Microsoft.Vbe.Interop;

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
            get { return System.IO.Path.GetFileNameWithoutExtension(name); }
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

        public void Save(string filename, Extensions extension)
        {
            this.ComObj.Application.DisplayAlerts = false;
            this.ComObj.SaveAs(filename, extension.ToExcel());
            this.ComObj.Application.DisplayAlerts = true;
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

        public void AddMacro(string name, string content, VbModuleType type)
        {
            Vi.VBComponent component = null;
            bool isNew = true;

            foreach (Vi.VBComponent comp in this.ComObj.VBProject.VBComponents)
            {
                if (comp.Name == name)
                {
                    isNew = false;
                    component = comp;
                }
            }
            if (isNew)
            {
                component = this.ComObj.VBProject.VBComponents.Add(type.ToExcel());
                component.Name = name;
            }

            component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines);
            component.CodeModule.AddFromString(content);
        }

        public void RunMacro(string name)
        {
            this.ComObj.Application.Run(name);
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Workbook | " + Name;
        }

        #endregion

    }
}
