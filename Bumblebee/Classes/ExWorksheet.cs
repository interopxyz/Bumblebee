using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

using XL = Microsoft.Office.Interop.Excel;

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
        }

        public virtual ExWorkbook Workbook
        {
            get { return new ExWorkbook((XL.Workbook)(this.ComObj.Parent)); }
        }

        #endregion

        #region methods

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

        #region data

        public void WriteData(List<ExRow> data, string source)
        {

            Tuple<int, int> location = Helper.GetCellLocation(source);

            int x = data[0].Values.Count;
            int y = data.Count;

            string[,] values = new string[y+1, x+0];

            for (int i = 0; i < data[0].Columns.Count; i++)
            {
                values[0, i] = data[0].Columns[i];
            }

            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    values[j + 1, i] = data[j].Values[i];
                }
            }

            string target = Helper.GetCellAddress(location.Item1 + x-2, location.Item2 + y-1);

            this.ComObj.Range[source, target].Value = values;

            for (int i = 0; i < data[0].Columns.Count; i++)
            {
                this.ComObj.Columns[location.Item2 + i].TextToColumns(Type.Missing, XL.XlTextParsingType.xlDelimited, XL.XlTextQualifier.xlTextQualifierNone);
            }

        }

        public void WriteData(List<ExColumn> data, string source)
        {

            Tuple<int, int> location = Helper.GetCellLocation(source);

            int x = data[0].Values.Count;
            int y = data.Count;

            string[,] values = new string[x + 1, y + 0];

            for (int i = 0; i < y; i++)
            {
                values[0, i] = data[i].Name;
            }

            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    values[i+1, j] = data[j].Values[i];
                }
            }

            string target = Helper.GetCellAddress(location.Item1 + y - 2, location.Item2 + x - 1);

            this.ComObj.Range[source, target].Value = values;

            for (int i = 0; i < data.Count; i++)
            {
                this.ComObj.Columns[location.Item2 + i].TextToColumns(Type.Missing, XL.XlTextParsingType.xlDelimited, XL.XlTextQualifier.xlTextQualifierNone);
            }

        }

        public void WriteData(List<List<GH_String>> data, string source)
        {

            Tuple<int, int> location = Helper.GetCellLocation(source);

            int y = data[0].Count;
            int x = data.Count;

            string[,] values = new string[y, x];

            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    values[j, i] = data[i][j].Value;
                }
            }

            string target = Helper.GetCellAddress(location.Item1 + x - 2, location.Item2 + y - 2);

            this.ComObj.Range[source, target].Value = values;

        }

        #endregion

        #region graphics

        public void ColorRange(string source, string target, Sd.Color color)
        {
            this.ComObj.Range[source, target].Interior.Color = color;
        }

        public void ColorCell(int row, int column, Sd.Color color)
        {
            this.ComObj.Cells[row, column].Interior.Color = color;
        }

        public void ColorCell(string cell, Sd.Color color)
        {
            Tuple<int, int> location = Helper.GetCellLocation(cell);
            this.ComObj.Cells[location.Item2, location.Item1].Interior.Color = color;
        }

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
