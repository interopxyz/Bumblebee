using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

using Rg = Rhino.Geometry;

using XL = Microsoft.Office.Interop.Excel;
using Grasshopper.Kernel;

namespace Bumblebee
{
    public class ExWorksheet
    {

        #region members

        public XL.Worksheet ComObj = null;

        protected string name = "";

        #endregion

        #region constructors

        public ExWorksheet()
        {
        }

        public ExWorksheet(ExRange range)
        {
            this.ComObj = range.ComObj.Worksheet;
            this.name = this.ComObj.Name;
        }

        public ExWorksheet(XL.Worksheet comObj)
        {
            this.ComObj = comObj;
            this.name = comObj.Name;
        }

        public ExWorksheet(ExWorksheet worksheet)
        {
            this.ComObj = worksheet.ComObj;
            this.name = worksheet.Name;
        }

        #endregion

        #region properties

        public virtual string Name
        {
            get { return name; }
            set 
            { 
                this.name = value;
                if (this.ComObj != null) this.ComObj.Name = this.name;
            }
        }

        public virtual ExWorkbook Workbook
        {
            get { return new ExWorkbook((XL.Workbook)(this.ComObj.Parent)); }
        }

        public virtual ExApp ParentApp
        {
            get { return new ExApp(this.ComObj.Application); }
        }

        #endregion

        #region methods

        public static ExWorksheet ActiveWorksheet()
        {
            ExApp app = new ExApp();
            return app.GetActiveWorksheet();
        }

        public static ExWorksheet ActiveWorksheet(string sheetName)
        {
            ExApp app = new ExApp();
            ExWorkbook book = app.GetActiveWorkbook();
            ExWorksheet sheet = book.GetWorksheet(sheetName);
            return sheet;
        }

        public void Freeze()
        {
            this.ComObj.Application.ScreenUpdating = false;
        }

        public void UnFreeze()
        {
            this.ComObj.Application.ScreenUpdating = true;
        }

        public void SetSheet(XL.Worksheet comObject)
        {
            this.ComObj = comObject;
            this.name = comObject.Name;
        }

        public void Activate()
        {
            this.ComObj.Activate();
        }

        public void Listen(GH_Component component, bool activate)
        {
            this.ComObj.Change -= (o) => { component.ExpireSolution(true); };
            if(activate) this.ComObj.Change += (o) => { component.ExpireSolution(true); };
        }

        public void ClearSheet()
        {
            ExRange range = this.GetUsedRange();
            range.ClearContent();
            range.ClearFormat();
        }

        #region data

        public ExRange WriteData(List<ExRow> data, ExCell source)
        {
            Dictionary<string, List<string>> values = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> formats = new Dictionary<string, List<string>>();

            foreach (ExRow row in data)
            {
                foreach (string name in row.Columns)
                {
                    if (!values.ContainsKey(name))
                    {
                    values.Add(name, new List<string>());
                    formats.Add(name, new List<string>());
                    }
                }
            }

            foreach (ExRow row in data)
            {
                foreach(string key in values.Keys)
                {
                    if (row.Columns.Contains(key))
                    {
                        int i = row.Columns.IndexOf(key);
                        values[key].Add(row.Values[i]);
                        formats[key].Add(row.Formats[i]);
                    }
                    else
                    {
                        values[key].Add("");
                        formats[key].Add("General");
                    }
                }
            }

            List<ExColumn> columns = new List<ExColumn>();
            foreach (string key in values.Keys)
            {
                columns.Add(new ExColumn(key, values[key], formats[key][0]));
            }

            return WriteData(columns, source);
        }

        public ExRange WriteData(List<ExColumn> data, ExCell source)
        {
            int x = data[0].Values.Count;
            int y = data.Count;

            List<bool> colNumeric = new List<bool>();

            Dictionary<int, List<ExColumn>> sets = new Dictionary<int, List<ExColumn>>();

            int step = 0;
            int count = 0;
            string type = "none";

            foreach (ExColumn col in data)
            {
                string current = "text";
                if (col.IsNumeric) current = "number";
                if (col.IsFormula) current = "formula";

                if (type != current)
                {
                    current = type;
                    count = step;
                    sets.Add(count, new List<ExColumn>());
                }

                sets[count].Add(col);
                step ++;
        }

            int c = source.Column;
            int r = source.Row;
            foreach (int key in sets.Keys)
            {
                string a = Helper.GetCellAddress(c + key, r);
                string b = Helper.GetCellAddress(c + key+sets[key].Count-1, r + x-1);

                if (sets[key][0].IsFormula)
                {
                    SetFunction(this.GetRange(new ExCell(a), new ExCell(b)).ComObj,sets[key].To2dTextArray());
                }
                else if (sets[key][0].IsNumeric)
                {
                    SetNumericData(this.GetRange(new ExCell(a), new ExCell(b)).ComObj, sets[key].To2dNumberArray());
                }
                else
                {
                    SetTextData(this.GetRange(new ExCell(a), new ExCell(b)).ComObj, sets[key].To2dTextArray());
                }
            }

            string target = Helper.GetCellAddress(source.Column + y - 1, source.Row + x);
            XL.Range rng = this.ComObj.Range[source.ToString(), target];

            for (int i = 0; i < data.Count; i++)
            {
                if(colNumeric[i]) this.ComObj.Range[new ExCell(source.Column+i,source.Row+1).ToString(), new ExCell(source.Column + i, source.Row+x).ToString()].TextToColumns(Type.Missing, XL.XlTextParsingType.xlDelimited, XL.XlTextQualifier.xlTextQualifierNone);
                this.ComObj.Columns[source.Column + i].NumberFormat = data[i].Format;
            }

            return new ExRange(rng);
        }

        public ExRange WriteData(List<List<GH_String>> data, ExCell source)
        {
            List<ExColumn> columns = new List<ExColumn>();
            foreach(List<GH_String> values in data)
            {
                List<string> text = new List<string>();
                foreach (GH_String val in values) text.Add(val.Value);
                columns.Add(new ExColumn(text));
            }

            return this.WriteData(columns, source);
        }

        protected void SetFunction(XL.Range rng, string[,] formulas, string format = "General")
        {
            rng.NumberFormat = format;
            string[] f = formulas.Flatten();
            int i = 0;
            foreach (XL.Range cell in rng.Cells)
            {
                cell.NumberFormat = "General";
                cell.Formula = f[i++];
            }
        }

        protected void SetTextData(XL.Range rng, string[,] values, string format = "@")
        {
            rng.NumberFormat = format;
            rng.Value2 = values;
        }

        protected void SetNumericData(XL.Range rng, double[,] values, string format = "0.00")
        {
            rng.NumberFormat = format;
            rng.Value2 = values;
        }

        #endregion

        #region range

        public ExRange GetRange(ExCell start, ExCell extent)
        {
            XL.Range rng = this.ComObj.Range[start.ToString(), extent.ToString()];
            return new ExRange(rng);
        }

        public ExRange GetRange(string range)
        {
            string[] cells = range.Split(':');
            ExCell start = new ExCell(cells[0]);
            ExCell extent = new ExCell(cells[0]);
            if (cells.Count() > 1) extent = new ExCell(cells[1]);

            XL.Range rng = this.ComObj.Range[start.ToString(), extent.ToString()];
            return new ExRange(rng);
        }

        public ExRange GetUsedRange()
        {
            return new ExRange(this.ComObj.UsedRange); ;
        }

        public string GetFirstUsedCell()
        {
            XL.Range rng = this.ComObj.UsedRange;

            int X = rng.Column;
            int Y = rng.Row;

            return Helper.GetCellAddress(X, Y);
        }

        public string GetLastUsedCell()
        {
            XL.Range rng = this.ComObj.UsedRange;

            int W = rng.Columns.Count;
            int X = rng.Columns[W].Column;
            int H = rng.Rows.Count;
            int Y = rng.Rows[H].Row;

            return Helper.GetCellAddress(X, Y);
        }

        #endregion

        #region objects

        public void AddPicture(string name, string fileName, double x, double y, double scale)
        {
            foreach(XL.Shape shp in this.ComObj.Shapes)
            {
                if (shp.Name == name) shp.Delete();
            }

            XL.Shape shape = this.ComObj.Shapes.AddPicture(fileName, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, (float)x, (float)y, 100, 100);
            shape.ScaleWidth((float)scale, Microsoft.Office.Core.MsoTriState.msoTrue);
            shape.ScaleHeight((float)scale, Microsoft.Office.Core.MsoTriState.msoTrue);
            shape.Name = name;
        }

        #endregion

        #region controls


        #endregion

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Worksheet | " + Name;
        }

        #endregion

    }
}
