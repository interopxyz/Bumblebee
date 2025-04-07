using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

using Rg = Rhino.Geometry;

using XL = Microsoft.Office.Interop.Excel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

namespace Bumblebee
{
    public class ExRange
    {

        #region members

        protected ExCell start = new ExCell();
        protected ExCell extent = new ExCell();

        public XL.Range ComObj = null;

        #endregion

        #region constructors

        public ExRange()
        {
        }

        public ExRange(XL.Range comObj)
        {
            this.ComObj = comObj;
            GetFirstCell();
            GetLastCell();
        }

        public ExRange(ExRange range)
        {
            this.ComObj = range.ComObj;
            this.start = range.Start;
            this.extent = range.Extent;
        }

        #endregion

        #region properties

        public virtual ExCell Start
        {
            get { return new ExCell(start); }
            set { start = new ExCell(value); }
        }

        public virtual ExCell Extent
        {
            get { return new ExCell(extent); }
            set { extent = new ExCell(value); }
        }

        public virtual ExWorksheet Worksheet
        {
            get { return new ExWorksheet(this.ComObj.Worksheet); }
        }

        public virtual ExWorkbook Workbook
        {
            get { return new ExWorkbook(this.Worksheet.Workbook); }
        }

        public virtual ExApp ParentApp
        {
            get { return new ExApp(this.ComObj.Application); }
        }

        public virtual bool IsSingle
        {
            get { return ((start.Column == extent.Column) & (start.Row == extent.Row)); }
        }

        public virtual string FontFamily
        {
            get { return this.ComObj.Font.Name; }
            set { this.ComObj.Font.Name = value; }
        }

        public virtual double FontSize
        {
            get { return this.ComObj.Font.Size; }
            set { this.ComObj.Font.Size = value; }
        }

        public virtual Sd.Color FontColor
        {
            get { return Sd.ColorTranslator.FromOle((int)this.ComObj.Font.Color); }
            set { this.ComObj.Font.Color = value; }
        }

        public virtual Justification FontJustification
        {
            set{
                this.ComObj.HorizontalAlignment = value.ToExcelHalign();
                this.ComObj.VerticalAlignment = value.ToExcelValign();
            }
            get
            {
                int align = 0;
                switch (this.ComObj.VerticalAlignment)
                {
                    case XL.XlVAlign.xlVAlignCenter:
                        align = 3;
                        break;
                    case XL.XlVAlign.xlVAlignTop:
                        align = 6;
                        break;
                }

                switch (this.ComObj.VerticalAlignment)
                {
                    case XL.XlHAlign.xlHAlignLeft:
                    case XL.XlHAlign.xlHAlignGeneral:
                        break;
                    case XL.XlHAlign.xlHAlignRight:
                        align += 2;
                        break;
                    default:
                        align += 1;
                        break;
                }
                return (Justification)align;
            }
        }

        public bool Bold
        {
            get { return this.ComObj.Font.Bold; }
            set { this.ComObj.Font.Bold = value; }
        }

        public bool Italic
        {
            get { return this.ComObj.Font.Italic; }
            set { this.ComObj.Font.Italic = value; }
        }

        public virtual BorderWeight Weight
        {
            get { return ((XL.XlBorderWeight)this.ComObj.Borders.Weight).ToBB(); }
        }

        public virtual LineType LineType
        {
            get { return ((XL.XlLineStyle)this.ComObj.Borders.LineStyle).ToBB(); }
        }

        public virtual Sd.Color BorderColor
        {
            get { return Sd.ColorTranslator.FromOle((int)this.ComObj.Borders.Color); }
        }

        public virtual Sd.Color Background
        {
            get { return Sd.ColorTranslator.FromOle((int)this.ComObj.Interior.Color); }
            set { this.ComObj.Interior.Color = value; }
        }

        public virtual int Width
        {
            get { return this.ComObj.Columns.ColumnWidth; }
            set { this.ComObj.Columns.ColumnWidth = value; }
        }

        public virtual int Height
        {
            get { return this.ComObj.Rows.RowHeight; }
            set { this.ComObj.Rows.RowHeight = value; }
        }

        #endregion

        #region methods

        #region data

        public GH_Structure<GH_String> ReadData()
        {
            GH_Structure<GH_String> data = new GH_Structure<GH_String>();

            int[] L = GetRangeArray();

            System.Array values = (System.Array)this.ComObj.Value2;

            if (values != null)
            {
                for (int i = 1; i < (L[3] - L[1] + 2); i++)
                {
                    GH_Path path = new GH_Path(i);
                    for (int j = 1; j < (L[2] - L[0] + 2); j++)
                    {
                        string val = string.Empty;
                        if (values.GetValue(i, j) != null) val = values.GetValue(i, j).ToString();
                        data.Append(new GH_String(val), path);
                    }
                }
            }

            return data;
        }

        #endregion

        #region Graphics

        public GH_Structure<GH_ObjectWrapper> ReadFillColors()
        {
            GH_Structure<GH_ObjectWrapper> ranges = new GH_Structure<GH_ObjectWrapper>();

            int Ax = this.Start.Column;
            int Ay = this.Start.Row;
            int Bx = this.Extent.Column;
            int By = this.Extent.Row;

            int k = 0;
            for (int i = Ax; i < Bx+1; i++)
                {
                    GH_Path path = new GH_Path(k);
                    for (int j = Ay; j < By+1; j++)
                    {
                    string address = new ExCell(i, j).Address;
                    ExRange rng = new ExRange(this.ComObj.Worksheet.Range[address,address]);
                    ranges.Append(new GH_ObjectWrapper(rng), path);
                    }
                k++;
                }

            return ranges;
        }

        #endregion

        #region Geometry

        protected int[] GetRangeArray()
        {
            return new int[] { this.Start.Column, this.Start.Row, this.Extent.Column, this.Extent.Row };
        }

        public Rg.Point3d GetMinPixel()
        {
            return new Rg.Point3d(this.ComObj.Left, this.ComObj.Top, 0);
        }

        public Rg.Point3d GetMaxPixel()
        {
            double X = this.ComObj.Left;
            double Y = this.ComObj.Top;

            double W = this.ComObj.Width;
            double H = this.ComObj.Height;

            return new Rg.Point3d(X + W, Y + H, 0);
        }

        #endregion

        #region Extract

        protected void GetFirstCell()
        {
            start= new ExCell(this.ComObj.Column, this.ComObj.Row, false, false);
        }

        protected ExRange GetFirstRange()
        {
            ExWorksheet sheet = new ExWorksheet(this.ComObj.Worksheet);
            return new ExRange(sheet.GetRange(start,start));
        }

        protected void GetLastCell()
        {
            extent= new ExCell(this.ComObj.Columns[this.ComObj.Columns.Count].Column, this.ComObj.Rows[this.ComObj.Rows.Count].Row, false, false);
        }

        public void ClearContent()
        {
            this.ComObj.ClearContents();
        }

        public void MergeCells()
        {
            this.ComObj.MergeCells = true;
        }

        public void UnMergeCells()
        {
            this.ComObj.UnMerge();
        }

        #endregion

        #region graphics

        public void ClearFormat()
        {
            this.ComObj.ClearFormats();
        }

        public void ClearBorders()
        {
            this.ComObj.Borders.LineStyle = XL.XlLineStyle.xlLineStyleNone;
        }

        public void ClearFill()
        {
            this.ComObj.Cells.Interior.Pattern = XL.XlPattern.xlPatternNone;
        }

        public void SetFont(string name, double size, Sd.Color color, Justification justification, bool bold, bool italic)
        {
            XL.Font font = this.ComObj.Font;

            font.Name = name;
            font.Size = size;
            font.Color = color;
            font.Bold = bold;
            font.Italic = italic;

            this.ComObj.HorizontalAlignment = justification.ToExcelHalign();
            this.ComObj.VerticalAlignment = justification.ToExcelValign();
        }

        public void SetBorder(Sd.Color color, BorderWeight weight, LineType type, HorizontalBorder horizontal, VerticalBorder vertical)
        {
            XL.Borders borders = this.ComObj.Borders;
            XL.Border left, right, top, bottom, betweenHorizontal, betweenVertical;

            switch (horizontal)
            {
                case HorizontalBorder.All:
                    bottom = borders[XL.XlBordersIndex.xlEdgeBottom];
                    bottom.Weight = weight.ToExcel();
                    bottom.Color = color;
                    bottom.LineStyle = type.ToExcel();

                    betweenHorizontal = borders[XL.XlBordersIndex.xlInsideHorizontal];
                    betweenHorizontal.Weight = weight.ToExcel();
                    betweenHorizontal.Color = color;
                    betweenHorizontal.LineStyle = type.ToExcel();

                    top = borders[XL.XlBordersIndex.xlEdgeTop];
                    top.Weight = weight.ToExcel();
                    top.Color = color;
                    top.LineStyle = type.ToExcel();
                    break;
                case HorizontalBorder.Between:
                    betweenHorizontal = borders[XL.XlBordersIndex.xlInsideHorizontal];
                    betweenHorizontal.Weight = weight.ToExcel();
                    betweenHorizontal.Color = color;
                    betweenHorizontal.LineStyle = type.ToExcel();
                    break;
                case HorizontalBorder.Both:
                    bottom = borders[XL.XlBordersIndex.xlEdgeBottom];
                    bottom.Weight = weight.ToExcel();
                    bottom.Color = color;
                    bottom.LineStyle = type.ToExcel();

                    top = borders[XL.XlBordersIndex.xlEdgeTop];
                    top.Weight = weight.ToExcel();
                    top.Color = color;
                    top.LineStyle = type.ToExcel();
                    break;
                case HorizontalBorder.Bottom:
                    bottom = borders[XL.XlBordersIndex.xlEdgeBottom];
                    bottom.Weight = weight.ToExcel();
                    bottom.Color = color;
                    bottom.LineStyle = type.ToExcel();
                    break;
                case HorizontalBorder.Top:
                    top = borders[XL.XlBordersIndex.xlEdgeTop];
                    top.Weight = weight.ToExcel();
                    top.Color = color;
                    top.LineStyle = type.ToExcel();
                    break;
            }

            switch (vertical)
            {
                case VerticalBorder.All:
                    left = borders[XL.XlBordersIndex.xlEdgeLeft];
                    left.Weight = weight.ToExcel();
                    left.Color = color;
                    left.LineStyle = type.ToExcel();

                    betweenVertical = borders[XL.XlBordersIndex.xlInsideVertical];
                    betweenVertical.Weight = weight.ToExcel();
                    betweenVertical.Color = color;
                    betweenVertical.LineStyle = type.ToExcel();

                    right = borders[XL.XlBordersIndex.xlEdgeRight];
                    right.Weight = weight.ToExcel();
                    right.Color = color;
                    right.LineStyle = type.ToExcel();
                    break;
                case VerticalBorder.Between:
                    betweenVertical = borders[XL.XlBordersIndex.xlInsideVertical];
                    betweenVertical.Weight = weight.ToExcel();
                    betweenVertical.Color = color;
                    betweenVertical.LineStyle = type.ToExcel();
                    break;
                case VerticalBorder.Both:
                    left = borders[XL.XlBordersIndex.xlEdgeLeft];
                    left.Weight = weight.ToExcel();
                    left.Color = color;
                    left.LineStyle = type.ToExcel();

                    right = borders[XL.XlBordersIndex.xlEdgeRight];
                    right.Weight = weight.ToExcel();
                    right.Color = color;
                    right.LineStyle = type.ToExcel();
                    break;
                case VerticalBorder.Left:
                    left = borders[XL.XlBordersIndex.xlEdgeLeft];
                    left.Weight = weight.ToExcel();
                    left.Color = color;
                    left.LineStyle = type.ToExcel();
                    break;
                case VerticalBorder.Right:
                    right = borders[XL.XlBordersIndex.xlEdgeRight];
                    right.Weight = weight.ToExcel();
                    right.Color = color;
                    right.LineStyle = type.ToExcel();
                    break;
            }

        }

        #endregion

        #region sparklines

        public void AddSparkLine(ExRange placement, Sd.Color color, double weight)
        {
            XL.SparklineGroup spark = placement.ComObj.SparklineGroups.Add(XL.XlSparkType.xlSparkLine, this.ToString());
            spark.SeriesColor.Color= color;
            spark.LineWeight = weight;
        }

        public void AddSparkColumn(ExRange placement, Sd.Color color)
        {
            XL.SparklineGroup spark = placement.ComObj.SparklineGroups.Add(XL.XlSparkType.xlSparkLine, this.ToString());
            spark.Type = XL.XlSparkType.xlSparkColumn;
            spark.SeriesColor.Color = color;
        }

        #endregion

        #region conditional

        public void ClearConditions()
        {
            this.ComObj.FormatConditions.Delete();
        }

        public void AddConditionalValue(ValueCondition condition, double value, Sd.Color color)
        {
            bool valid = true;

            switch (condition)
            {
                case ValueCondition.Greater:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlCellValue, XL.XlFormatConditionOperator.xlGreater, Formula1: value);
                    break;
                case ValueCondition.GreaterEqual:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlCellValue, XL.XlFormatConditionOperator.xlGreaterEqual, Formula1: value);
                    break;
                case ValueCondition.Less:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlCellValue, XL.XlFormatConditionOperator.xlLess, Formula1: value);
                    break;
                case ValueCondition.LessEqual:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlCellValue, XL.XlFormatConditionOperator.xlLessEqual, Formula1: value);
                    break;
                case ValueCondition.Equal:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlCellValue, XL.XlFormatConditionOperator.xlEqual, Formula1: value);
                    break;
                case ValueCondition.NotEqual:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlCellValue, XL.XlFormatConditionOperator.xlNotEqual, Formula1: value);
                    break;
                default:
                    valid = false;
                    break;
            }
            if (valid) this.ComObj.FormatConditions[this.ComObj.FormatConditions.Count].Interior.Color = color;

        }

        public void AddConditionalAverage(AverageCondition condition, Sd.Color color)
        {
            bool valid = true;

            switch (condition)
            {
                case AverageCondition.AboveAverage:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlAboveAverageCondition, XL.XlAboveBelow.xlAboveAverage);
                    break;
                case AverageCondition.AboveDeviation:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlAboveAverageCondition, XL.XlAboveBelow.xlAboveStdDev);
                    break;
                case AverageCondition.AboveEqualAverage:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlAboveAverageCondition, XL.XlAboveBelow.xlEqualAboveAverage);
                    break;
                case AverageCondition.BelowAverage:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlAboveAverageCondition, XL.XlAboveBelow.xlBelowAverage);
                    break;
                case AverageCondition.BelowDeviation:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlAboveAverageCondition, XL.XlAboveBelow.xlBelowStdDev);
                    break;
                case AverageCondition.BelowEqualAverage:
                    this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlAboveAverageCondition, XL.XlAboveBelow.xlEqualBelowAverage);
                    break;
                default:
                    valid = false;
                    break;
            }
            if (valid) this.ComObj.FormatConditions[this.ComObj.FormatConditions.Count].Interior.Color = color;

        }

        public void AddConditionalBetween(double low, double high, Sd.Color color, bool flip=false)
        {
            if (flip)
            {
                this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlCellValue, XL.XlFormatConditionOperator.xlNotBetween, Formula1: low,Formula2:high);
            }
            else
            {
                this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlCellValue, XL.XlFormatConditionOperator.xlBetween, Formula1: low, Formula2: high);
            }
            this.ComObj.FormatConditions[this.ComObj.FormatConditions.Count].Interior.Color = color;
        }

        public void AddConditionalUnique(Sd.Color color, bool flip = false)
        {
            XL.UniqueValues unique = this.ComObj.FormatConditions.AddUniqueValues();
            if (flip)
            {
                unique.DupeUnique = XL.XlDupeUnique.xlDuplicate;
            }
            else
            {
                unique.DupeUnique = XL.XlDupeUnique.xlUnique;
            }
            unique.Interior.Color = color;
        }

        public void AddConditionalTopCount(int count, Sd.Color color, bool flip = false)
        {
            XL.Top10 top = this.ComObj.FormatConditions.AddTop10();
            top.Percent = false;
            top.Rank = count;
            if (flip)
            {
                top.TopBottom = XL.XlTopBottom.xlTop10Bottom;
            }
            else
            {
                top.TopBottom = XL.XlTopBottom.xlTop10Top;
            }
            top.Interior.Color = color;
        }

        public void AddConditionalTopPercent(double percent, Sd.Color color, bool flip = false)
        {
            percent = Math.Min(percent, 1.0);
            percent = Math.Max(percent, 0.0);
            XL.Top10 top = this.ComObj.FormatConditions.AddTop10();
            top.Percent = true;
            top.Rank = (int)(percent*100.0);
            if (flip)
            {
                top.TopBottom = XL.XlTopBottom.xlTop10Bottom;
            }
            else
            {
                top.TopBottom = XL.XlTopBottom.xlTop10Top;
            }
            top.Interior.Color = color;
        }

        public void AddConditionalBlanks(Sd.Color color, bool flip = false)
        {

            if (flip)
            {
                this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlNoBlanksCondition);
            }
            else
            {
                this.ComObj.FormatConditions.Add(XL.XlFormatConditionType.xlBlanksCondition);
            }
            this.ComObj.FormatConditions[this.ComObj.FormatConditions.Count].Interior.Color = color;
        }

        public void AddConditionalScale(Sd.Color first, Sd.Color second, Sd.Color third, double mid = 0.5)
        {
            mid = Math.Min(mid, 1.0);
            mid = Math.Max(mid, 0.0);
            XL.ColorScale scale = this.ComObj.FormatConditions.AddColorScale(3);

            scale.ColorScaleCriteria[1].Type = XL.XlConditionValueTypes.xlConditionValueLowestValue;
            scale.ColorScaleCriteria[1].FormatColor.Color = first;

            scale.ColorScaleCriteria[2].Type = XL.XlConditionValueTypes.xlConditionValuePercentile;
            scale.ColorScaleCriteria[2].Value = (int)(mid*100.0);
            scale.ColorScaleCriteria[2].FormatColor.Color = second;

            scale.ColorScaleCriteria[3].Type = XL.XlConditionValueTypes.xlConditionValueHighestValue;
            scale.ColorScaleCriteria[3].FormatColor.Color = third;
        }

        public void AddConditionalScale(Sd.Color first, Sd.Color second)
        {
            XL.ColorScale scale = this.ComObj.FormatConditions.AddColorScale(2);
            int count = this.ComObj.FormatConditions.Count;

            scale.ColorScaleCriteria[1].Type = XL.XlConditionValueTypes.xlConditionValueLowestValue;
            scale.ColorScaleCriteria[1].FormatColor.Color = first;

            scale.ColorScaleCriteria[2].Type = XL.XlConditionValueTypes.xlConditionValueHighestValue;
            scale.ColorScaleCriteria[2].FormatColor.Color = second;
        }

        public void AddConditionalBar(Sd.Color color, bool gradient)
        {
            XL.Databar bar = this.ComObj.FormatConditions.AddDatabar();

            bar.BarColor.Color = color;
            if (gradient)
            {
                bar.BarFillType = XL.XlDataBarFillType.xlDataBarFillGradient;
            }
            else
            {
                bar.BarFillType = XL.XlDataBarFillType.xlDataBarFillSolid;
            }

        }

        #endregion

        #endregion

        #region overrides

        public override string ToString()
        {
            return this.start.Address+":"+this.extent.Address;
        }

        #endregion

    }
}
