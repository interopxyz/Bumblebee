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

            string[,] values = new string[x+1, y];
            double[,] numbers = new double[x+1, y];
            bool isNumeric = true;
            List<bool> colNumeric = new List<bool>();

            Dictionary<int, List<ExColumn>> sets = new Dictionary<int, List<ExColumn>>();

            int step = 0;
            int count = 0;
            string type = "text";
            sets.Add(0, new List<ExColumn>());
            foreach(ExColumn col in data)
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
                step += 1;
        }

        double num;
            for (int i = 0; i < y; i++)
            {
                values[0, i] = data[i].Name;
                if (double.TryParse(values[0, i], out num)) numbers[0, i] = num; else isNumeric = false;
                colNumeric.Add(true);
            }

            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    values[i + 1, j] = data[j].Values[i];
                    if (double.TryParse(values[i + 1, j], out num))
                    {
                        numbers[i + 1, j] = num;
                    }
                    else
                    {
                        isNumeric = false;
                        colNumeric[j] = false;
                    }
                    
                }
            }

            string target = Helper.GetCellAddress(source.Column + y - 1, source.Row + x);

            XL.Range rng = this.ComObj.Range[source.ToString(), target];

            if (isNumeric) SetNumericData(rng, numbers);
            if (!isNumeric) SetTextData(rng, values);

            for (int i = 0; i < data.Count; i++)
            {
                if(colNumeric[i]) this.ComObj.Range[new ExCell(source.Column+i,source.Row+1).ToString(), new ExCell(source.Column + i, source.Row+x).ToString()].TextToColumns(Type.Missing, XL.XlTextParsingType.xlDelimited, XL.XlTextQualifier.xlTextQualifierNone);
                    this.ComObj.Columns[source.Column + i].NumberFormat = data[i].Format;
            }

            return new ExRange(rng);
        }

        public ExRange WriteData(List<List<GH_String>> data, ExCell source)
        {
            int y = data[0].Count;
            int x = data.Count;

            string[,] values = new string[y, x];
            double[,] numbers = new double[y, x];
            bool isNumeric = true;
            bool isFunction = true;

            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    values[j, i] = data[i][j].Value;
                    if (double.TryParse(values[j, i], out double num)) numbers[j, i] = num; else isNumeric = false;
                    if (!(values[j, i].ToCharArray()[0].Equals('='))) isFunction = false;
                }
            }

            string target = Helper.GetCellAddress(source.Column + x - 1, source.Row + y - 1);
            XL.Range rng = this.ComObj.Range[source.ToString(), target];

            if (isNumeric)
            {
                SetNumericData(rng, numbers);
            }
            else if (isFunction)
            {
                SetFunction(rng, values);
            }
            else
            {
                SetTextData(rng, values);
            }

            return GetRange(source, new ExCell(target));
        }

        protected void SetFunction(XL.Range rng, string[,] formulas)
        {
            string[] f = formulas.Flatten();
            int i = 0;
            foreach (XL.Range cell in rng.Cells)
            {
                cell.Formula = f[i++];
            }
        }

        protected void SetTextData(XL.Range rng, string[,] values)
        {
            rng.NumberFormat = "@";
            rng.Value2 = values;
        }

        protected void SetNumericData(XL.Range rng, double[,] values)
        {
            rng.NumberFormat = "0.00";
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
