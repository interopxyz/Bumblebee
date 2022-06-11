using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Rg = Rhino.Geometry;
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

        public string WriteData(List<ExRow> data, string source)
        {

            Tuple<int, int> location = Helper.GetCellLocation(source);

            int x = data[0].Values.Count;
            int y = data.Count;

            string[,] values = new string[y + 1, x + 0];

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

            string target = Helper.GetCellAddress(location.Item1 + x - 1, location.Item2 + y);

            this.ComObj.Range[source, target].Value = values;

            for (int i = 0; i < data[0].Columns.Count; i++)
            {
                this.ComObj.Columns[location.Item2 + i].TextToColumns(Type.Missing, XL.XlTextParsingType.xlDelimited, XL.XlTextQualifier.xlTextQualifierNone);
            }

            return target;
        }

        public string WriteData(List<ExColumn> data, string source)
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
                    values[i + 1, j] = data[j].Values[i];
                }
            }

            string target = Helper.GetCellAddress(location.Item1 + y - 1, location.Item2 + x);

            this.ComObj.Range[source, target].Value = values;

            for (int i = 0; i < data.Count; i++)
            {
                this.ComObj.Columns[location.Item2 + i].TextToColumns(Type.Missing, XL.XlTextParsingType.xlDelimited, XL.XlTextQualifier.xlTextQualifierNone);
            }

            return target;
        }

        public string WriteData(List<List<GH_String>> data, string source)
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

            string target = Helper.GetCellAddress(location.Item1 + x - 1, location.Item2 + y - 1);

            this.ComObj.Range[source, target].Value = values;

            return target;
        }

        #endregion

        #region range

        public void ResizeRangeCells(string source, string target, int width, int height)
        {

            this.ComObj.Range[source, target].Columns.ColumnWidth = width;
            this.ComObj.Range[source, target].Rows.RowHeight = height;
        }

        public void ClearContent(string source, string target)
        {
            this.ComObj.Range[source, target].ClearContents();
        }

        public void ClearFormat(string source, string target)
        {
            this.ComObj.Range[source, target].ClearFormats();
        }

        public void MergeCells(string source, string target)
        {
            this.ComObj.Range[source, target].MergeCells = true;
        }

        public void UnMergeCells(string source, string target)
        {
            this.ComObj.Range[source, target].UnMerge();
        }

        public Rg.Point3d GetRangeMinPixel(string source, string target)
        {
            double X = this.ComObj.Range[source, target].Left;
            double Y = this.ComObj.Range[source, target].Top;

            return new Rg.Point3d(X, Y, 0);
        }

        public Rg.Point3d GetRangeMaxPixel(string source, string target)
        {
            double X = this.ComObj.Range[source, target].Left;
            double Y = this.ComObj.Range[source, target].Top;

            double W = this.ComObj.Range[source, target].Width;
            double H = this.ComObj.Range[source, target].Height;

            return new Rg.Point3d(X + W, Y + H, 0);
        }

        public string GetFirstUsedCell()
        {
            int X = this.ComObj.UsedRange.Column;
            int Y = this.ComObj.UsedRange.Row;

            return Helper.GetCellAddress(X, Y);
        }

        public string GetLastUsedCell()
        {
            int W = this.ComObj.UsedRange.Columns.Count;
            int X = this.ComObj.UsedRange.Columns[W].Column;
            int H = this.ComObj.UsedRange.Rows.Count;
            int Y = this.ComObj.UsedRange.Rows[H].Row;

            return Helper.GetCellAddress(X, Y);
        }

        #endregion

        #region graphics

        public void RangeFont(string source, string target, string name, double size, Sd.Color color, ExApp.Justification justification, bool bold, bool italic)
        {
            XL.Range range = this.ComObj.Range[source, target];
            XL.Font font = range.Font;

            font.Name = name;
            font.Size = size;
            font.Color = color;
            font.Bold = bold;
            font.Italic = italic;

            range.HorizontalAlignment = justification.ToExcelHalign();
            range.VerticalAlignment = justification.ToExcelValign();

        }

        public void RangeColor(string source, string target, Sd.Color color)
        {
            this.ComObj.Range[source, target].Interior.Color = color;
        }

        public void RangeBorder(string source, string target, Sd.Color color, ExApp.BorderWeight weight, ExApp.LineType type, ExApp.HorizontalBorder horizontal, ExApp.VerticalBorder vertical)
        {
            XL.Borders borders = this.ComObj.Range[source, target].Borders;
            XL.Border left, right, top, bottom;

            switch (horizontal)
            {
                case ExApp.HorizontalBorder.Both:
                    bottom = borders[XL.XlBordersIndex.xlEdgeBottom];
                    bottom.Weight = weight.ToExcel();
                    bottom.Color = color;
                    bottom.LineStyle = type.ToExcel();

                    top = borders[XL.XlBordersIndex.xlEdgeTop];
                    top.Weight = weight.ToExcel();
                    top.Color = color;
                    top.LineStyle = type.ToExcel();
                    break;
                case ExApp.HorizontalBorder.Bottom:
                    bottom = borders[XL.XlBordersIndex.xlEdgeBottom];
                    bottom.Weight = weight.ToExcel();
                    bottom.Color = color;
                    bottom.LineStyle = type.ToExcel();
                    break;
                case ExApp.HorizontalBorder.Top:
                    top = borders[XL.XlBordersIndex.xlEdgeTop];
                    top.Weight = weight.ToExcel();
                    top.Color = color;
                    top.LineStyle = type.ToExcel();
                    break;
            }

            switch (vertical)
            {
                case ExApp.VerticalBorder.Both:
                    left = borders[XL.XlBordersIndex.xlEdgeLeft];
                    left.Weight = weight.ToExcel();
                    left.Color = color;
                    left.LineStyle = type.ToExcel();

                    right = borders[XL.XlBordersIndex.xlEdgeRight];
                    right.Weight = weight.ToExcel();
                    right.Color = color;
                    right.LineStyle = type.ToExcel();
                    break;
                case ExApp.VerticalBorder.Left:
                    left = borders[XL.XlBordersIndex.xlEdgeLeft];
                    left.Weight = weight.ToExcel();
                    left.Color = color;
                    left.LineStyle = type.ToExcel();
                    break;
                case ExApp.VerticalBorder.Right:
                    right = borders[XL.XlBordersIndex.xlEdgeRight];
                    right.Weight = weight.ToExcel();
                    right.Color = color;
                    right.LineStyle = type.ToExcel();
                    break;
            }

        }

        #endregion

        #region objects

        public void AddPicture(string fileName, double x, double y, double scale)
        {
            XL.Shape shape = this.ComObj.Shapes.AddPicture(fileName, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, (float)x, (float)y, 100, 100);
            shape.ScaleWidth((float)scale, Microsoft.Office.Core.MsoTriState.msoTrue);
            shape.ScaleHeight((float)scale, Microsoft.Office.Core.MsoTriState.msoTrue);

        }

        public void AddSparkline(string source, string target, string placement)
        {
            XL.Range range = this.ComObj.Range[source, target];
            range.SparklineGroups.Add(XL.XlSparkType.xlSparkColumn,placement+":"+placement);
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
