using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public void WriteData(List<ExData> data, string source)
        {

            Tuple<int, int> location = Helper.GetCellLocation(source);

            int x = data[0].Values.Count;
            int y = data.Count + 1;

            string[,] values = new string[y, x];

            for (int i = 0; i < data[0].Columns.Count; i++)
            {
                values[0, i] = data[0].Columns[i];
            }

                for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Columns.Count; j++)
                {
                    values[i+1, j] = data[i].Values[j];
                }
            }

            string target = Helper.GetCellAddress(location.Item1 + x, location.Item2 + y-2);

            this.ComObj.Range[source, target].Value = values;

            for (int i = 0; i < data[0].Columns.Count; i++)
            {

                this.ComObj.Columns[location.Item2+i].TextToColumns(Type.Missing, XL.XlTextParsingType.xlDelimited, XL.XlTextQualifier.xlTextQualifierNone);

            //    string start = Helper.GetCellAddress(location.Item1 + i, location.Item2);
            //    string end = Helper.GetCellAddress(location.Item1 + i, location.Item2 + y);
            //    string format = "\"" + data[0].Formats[i] + "\"";
            //this.ComObj.Cells[start, end].NumberFormat = format;
            }
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
