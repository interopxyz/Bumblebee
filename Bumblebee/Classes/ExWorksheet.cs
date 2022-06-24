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

        public string WriteData(List<ExRow> data, ExCell source)
        {
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

            string target = Helper.GetCellAddress(source.Column+ x - 1, source.Row + y);

            this.ComObj.Range[source.ToString(), target].Value = values;

            for (int i = 0; i < data[0].Columns.Count; i++)
            {
                this.ComObj.Columns[source.Row + i].TextToColumns(Type.Missing, XL.XlTextParsingType.xlDelimited, XL.XlTextQualifier.xlTextQualifierNone);
            }

            return target;
        }

        public string WriteData(List<ExColumn> data, ExCell source)
        {
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

            string target = Helper.GetCellAddress(source.Column + y - 1, source.Row + x);

            this.ComObj.Range[source.ToString(), target].Value = values;

            for (int i = 0; i < data.Count; i++)
            {
                this.ComObj.Columns[source.Row + i].TextToColumns(Type.Missing, XL.XlTextParsingType.xlDelimited, XL.XlTextQualifier.xlTextQualifierNone);
            }

            return target;
        }

        public ExRange WriteData(List<List<GH_String>> data, ExCell source)
        {
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

            string target = Helper.GetCellAddress(source.Column + x - 1, source.Row + y - 1);

            this.ComObj.Range[source.ToString(), target].Value = values;

            return new ExRange(source,new ExCell(target));
        }

        #endregion

        #region range

        public void RangeWidth(ExRange range, int width)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];
            rng.Columns.ColumnWidth = width;
        }

        public void RangeHeight(ExRange range, int height)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];
            rng.Rows.RowHeight = height;
        }

        public void ClearContent(ExRange range)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];
            rng.ClearContents();
        }

        public void ClearFormat(ExRange range)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];
            rng.ClearFormats();
        }

        public void MergeCells(ExRange range)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];
            rng.MergeCells = true;
        }

        public void UnMergeCells(ExRange range)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];
            rng.UnMerge();
        }

        public Rg.Point3d GetRangeMinPixel(ExRange range)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];

            double X = rng.Left;
            double Y = rng.Top;

            return new Rg.Point3d(X, Y, 0);
        }

        public Rg.Point3d GetRangeMaxPixel(ExRange range)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];

            double X = rng.Left;
            double Y = rng.Top;

            double W = rng.Width;
            double H = rng.Height;

            return new Rg.Point3d(X + W, Y + H, 0);
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

        #region graphics

        public void RangeFont(ExRange range, string name, double size, Sd.Color color, ExApp.Justification justification, bool bold, bool italic)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];
            XL.Font font = rng.Font;

            font.Name = name;
            font.Size = size;
            font.Color = color;
            font.Bold = bold;
            font.Italic = italic;

            rng.HorizontalAlignment = justification.ToExcelHalign();
            rng.VerticalAlignment = justification.ToExcelValign();
        }

        public void RangeColor(ExRange range, Sd.Color color)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];
            rng.Interior.Color = color;
        }

        public void RangeBorder(ExRange range, Sd.Color color, ExApp.BorderWeight weight, ExApp.LineType type, ExApp.HorizontalBorder horizontal, ExApp.VerticalBorder vertical)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];
            XL.Borders borders = rng.Borders;
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

        public void AddSparkline(ExRange range, string placement)
        {
            XL.Range rng = this.ComObj.Range[range.Start.ToString(), range.Extent.ToString()];
            rng.SparklineGroups.Add(XL.XlSparkType.xlSparkColumn,placement+":"+placement);
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
