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
    public class ExChart
    {
        #region members

        public XL.ChartObject ComObj = null;
        protected ExRange range = null;

        #endregion

        #region constructors

        public ExChart(string name, ExRange range, bool flip, Rg.Rectangle3d boundary)
        {
            this.range = new ExRange(range);

            bool isNew = true;
            foreach(XL.ChartObject obj in range.ComObj.Worksheet.ChartObjects())
            {
                if(name == obj.Name)
                {
                    this.ComObj = obj;
                    isNew = false;
                    break;
                }
            }

            if(isNew)this.ComObj = range.ComObj.Worksheet.ChartObjects().Add(boundary.Corner(0).X, boundary.Corner(0).Y, boundary.Width, boundary.Height);
            this.ComObj.Chart.SetSourceData(range.ComObj);
            if (flip)
            {
                this.ComObj.Chart.PlotBy = XL.XlRowCol.xlColumns;
            }
            else
            {
                this.ComObj.Chart.PlotBy = XL.XlRowCol.xlRows;
            }
            this.ComObj.Chart.HasLegend = false;
            foreach(XL.Series series in this.ComObj.Chart.SeriesCollection())
            {
                series.MarkerStyle = XL.XlMarkerStyle.xlMarkerStyleDot;
            }
            this.ComObj.Name = name;
        }

        public ExChart(ExChart chart)
        {
            this.ComObj = chart.ComObj;
            this.range = new ExRange(chart.range);
        }

        #endregion

        #region properties



        #endregion

        #region methods

        public void SetFillColors(bool bySeries, List<Sd.Color> values)
        {
            int i = 0;
            int c = values.Count - 1;
            range.Worksheet.Freeze();
            if (bySeries)
            {
                foreach (XL.Series series in this.ComObj.Chart.SeriesCollection())
                {
                    int j = i;
                    if (i > c) j = c;
                    series.Interior.Color = values[j];
                    i++;
                }
            }
            else
            {
                foreach (XL.Series series in this.ComObj.Chart.SeriesCollection())
                {
                    foreach (XL.Point point in series.Points())
                    {
                        int j = i;
                        if (i > c) j = c;
                        point.Interior.Color = values[j];
                        i++;
                    }
                }
            }
            range.Worksheet.UnFreeze();

        }

        public void SetStrokeColors(bool bySeries, List<Sd.Color> values)
        {
            int i = 0;
            int c = values.Count - 1;

            range.Worksheet.Freeze();
            if (bySeries)
            {
                foreach (XL.Series series in this.ComObj.Chart.SeriesCollection())
                {
                    int j = i;
                    if (i > c) j = c;
                    series.Border.Color = values[j];
                    i++;
                }
            }
            else
            {
                foreach (XL.Series series in this.ComObj.Chart.SeriesCollection())
                {
                    foreach (XL.Point point in series.Points())
                    {
                        int j = i;
                        if (i > c) j = c;
                        point.Border.Color = values[j];
                        i++;
                    }
                }
            }
            range.Worksheet.UnFreeze();

        }

        public void SetStrokeWeights(bool bySeries, List<int> values)
        {
            int i = 0;
            int c = values.Count - 1;

            range.Worksheet.Freeze();
            if (bySeries)
            {
                foreach (XL.Series series in this.ComObj.Chart.SeriesCollection())
                {
                    int j = i;
                    if (i > c) j = c;
                    if (values[j] > 3)
                    {
                        series.Border.LineStyle = XL.XlLineStyle.xlLineStyleNone;
                    }
                    else
                    {
                    series.Border.Weight = ((BorderWeight)values[j]).ToExcel();
                    }
                    i++;
                }
            }
            else
            {
                foreach (XL.Series series in this.ComObj.Chart.SeriesCollection())
                {
                    foreach (XL.Point point in series.Points())
                    {
                        int j = i;
                        if (i > c) j = c;
                        if (values[j] > 3)
                        {
                            point.Border.LineStyle = XL.XlLineStyle.xlLineStyleNone;
                        }
                        else
                        {
                            point.Border.Weight = ((BorderWeight)values[j]).ToExcel();
                        }
                        i++;
                    }
                }
            }
            range.Worksheet.UnFreeze();

        }

        public void SetFontNames(bool bySeries, List<double> values)
        {
            int i = 0;
            int c = values.Count - 1;

            if (bySeries)
            {
                foreach (XL.Series series in this.ComObj.Chart.SeriesCollection())
                {
                    int j = i;
                    if (i > c) j = c;
                    series.Border.Weight = values[j];
                    i++;
                }
            }
            else
            {
                foreach (XL.Series series in this.ComObj.Chart.SeriesCollection())
                {
                    foreach (XL.Point point in series.Points())
                    {
                        int j = i;
                        if (i > c) j = c;
                        point.Border.Weight = values[j];
                        i++;
                    }
                }
            }
        }

        public void SetTitle(string title)
        {
                this.ComObj.Chart.HasTitle = true;
            this.ComObj.Chart.ChartTitle.Text = title;
            if (title == "") this.ComObj.Chart.HasTitle = false;
        }

        public void SetLabel(LabelType label)
        {
            switch (label)
            {
                default:
                    this.ComObj.Chart.ApplyDataLabels(XL.XlDataLabelsType.xlDataLabelsShowNone);
                    break;
                case LabelType.Category:
                    this.ComObj.Chart.ApplyDataLabels(XL.XlDataLabelsType.xlDataLabelsShowLabel);
                    break;
                case LabelType.Value:
                    this.ComObj.Chart.ApplyDataLabels(XL.XlDataLabelsType.xlDataLabelsShowValue);
                    break;
            }
        }

        public void SetLegend(LegendLocations legend)
        {
            switch (legend)
            {
                case LegendLocations.None:
                    this.ComObj.Chart.HasLegend = false;
                    break;
                case LegendLocations.Bottom:
                    this.ComObj.Chart.HasLegend = true;
                    this.ComObj.Chart.Legend.Position = XL.XlLegendPosition.xlLegendPositionBottom;
                    break;
                case LegendLocations.Left:
                    this.ComObj.Chart.HasLegend = true;
                    this.ComObj.Chart.Legend.Position = XL.XlLegendPosition.xlLegendPositionLeft;
                    break;
                case LegendLocations.Right:
                    this.ComObj.Chart.HasLegend = true;
                    this.ComObj.Chart.Legend.Position = XL.XlLegendPosition.xlLegendPositionRight;
                    break;
                case LegendLocations.Top:
                    this.ComObj.Chart.HasLegend = true;
                    this.ComObj.Chart.Legend.Position = XL.XlLegendPosition.xlLegendPositionTop;
                    break;
            }
        }

        public void SetAxisX(string name)
        {
            XL.Axis axis = this.ComObj.Chart.Axes(XL.XlAxisType.xlValue, XL.XlAxisGroup.xlPrimary);
            axis.HasTitle = true;
            axis.AxisTitle.Text = name;

            if (name== "") axis.HasTitle = false;
        }

        public void SetAxisY(string name)
        {
            XL.Axis axis = this.ComObj.Chart.Axes(XL.XlAxisType.xlCategory, XL.XlAxisGroup.xlPrimary);
            axis.HasTitle = true;
            axis.AxisTitle.Text = name;

            if (name == "") axis.HasTitle = false;
        }

        public void SetGridX(GridType type)
        {
            XL.Axis axis = this.ComObj.Chart.Axes(XL.XlAxisType.xlValue, XL.XlAxisGroup.xlPrimary);
            switch (type)
            {
                case GridType.None:
                    axis.HasMajorGridlines = false;
                    axis.HasMinorGridlines = false;
                    break;
                case GridType.Primary:
                    axis.HasMajorGridlines = true;
                    axis.HasMinorGridlines = false;
                    break;
                case GridType.All:
                    axis.HasMajorGridlines = true;
                    axis.HasMinorGridlines = true;
                    break;
            }

        }

        public void SetGridY(GridType type)
        {
            XL.Axis axis = this.ComObj.Chart.Axes(XL.XlAxisType.xlCategory, XL.XlAxisGroup.xlPrimary);
            switch (type)
            {
                case GridType.None:
                    axis.HasMajorGridlines = false;
                    axis.HasMinorGridlines = false;
                    break;
                case GridType.Primary:
                    axis.HasMajorGridlines = true;
                    axis.HasMinorGridlines = false;
                    break;
                case GridType.All:
                    axis.HasMajorGridlines = true;
                    axis.HasMinorGridlines = true;
                    break;
            }
        }

        public void SetBarChart(BarChartType chartType, ChartFill fillType)
        {
            int type = 57;
            switch (chartType)
            {
                case BarChartType.Basic:
                    type = 57;
                    break;
                case BarChartType.Box:
                    type = 60;
                    break;
                case BarChartType.Pyramid:
                    type = 109;
                    break;
                case BarChartType.Cylinder:
                    type = 95;
                    break;
                case BarChartType.Cone:
                    type = 102;
                    break;
            }

            type += (int)fillType;


            this.ComObj.Chart.ChartType = (XL.XlChartType)type;
        }

        public void SetColumnChart(BarChartType chartType, ChartFill fillType)
        {
            int type = 51;
            switch (chartType)
            {
                case BarChartType.Basic:
                    type = 51;
                    break;
                case BarChartType.Box:
                    type = 54;
                    break;
                case BarChartType.Pyramid:
                    type = 106;
                    break;
                case BarChartType.Cylinder:
                    type = 92;
                    break;
                case BarChartType.Cone:
                    type = 99;
                    break;
            }

            type += (int)fillType;


            this.ComObj.Chart.ChartType = (XL.XlChartType)type;
        }

        public void SetRadialChart(RadialChartType chartType)
        {
            XL.XlChartType type = XL.XlChartType.xlPie;
            switch (chartType)
            {
                case RadialChartType.Pie:
                    type = XL.XlChartType.xlPie;
                    break;
                case RadialChartType.Donut:
                    type = XL.XlChartType.xlDoughnut;
                    break;
                case RadialChartType.Radar:
                    type = XL.XlChartType.xlRadar;
                    break;
                case RadialChartType.Pie3D:
                    type = XL.XlChartType.xl3DPie;
                    break;
                case RadialChartType.RadarFilled:
                    type = XL.XlChartType.xlRadarFilled;
                    break;
            }

            this.ComObj.Chart.ChartType = type;
        }

        public void SetScatterChart(ScatterChartType chartType)
        {
            XL.XlChartType type = XL.XlChartType.xlXYScatter;
            switch (chartType)
            {
                case ScatterChartType.Scatter:
                    type = XL.XlChartType.xlXYScatter;
                    break;
                case ScatterChartType.ScatterLines:
                    type = XL.XlChartType.xlXYScatterLines;
                    break;
                case ScatterChartType.ScatterSmooth:
                    type = XL.XlChartType.xlXYScatterSmooth;
                    break;
                case ScatterChartType.Bubble:
                    type = XL.XlChartType.xlBubble;
                    break;
                case ScatterChartType.Bubble3D:
                    type = XL.XlChartType.xlBubble3DEffect;
                    break;
            }

            this.ComObj.Chart.ChartType = type;
        }

        public void SetSurfaceChart(SurfaceChartType chartType)
        {
            XL.XlChartType type = XL.XlChartType.xlSurface;
            switch (chartType)
            {
                case SurfaceChartType.Surface:
                    type = XL.XlChartType.xlSurface;
                    break;
                case SurfaceChartType.SurfaceTop:
                    type = XL.XlChartType.xlSurfaceTopView;
                    break;
                case SurfaceChartType.SurfaceWireframe:
                    type = XL.XlChartType.xlSurfaceWireframe;
                    break;
                case SurfaceChartType.SurfaceWireframeTop:
                    type = XL.XlChartType.xlSurfaceTopViewWireframe;
                    break;
            }

            this.ComObj.Chart.ChartType = type;
        }

        public void SetLineChart(LineChartType chartType, ChartFill fillType)
        {
            XL.XlChartType type = XL.XlChartType.xlLine;
            switch (chartType)
            {
                case LineChartType.Line:
                    switch (fillType)
                    {
                        case ChartFill.Stack:
                            type = XL.XlChartType.xlLineStacked;
                            break;
                        case ChartFill.Fill:
                            type = XL.XlChartType.xlLineStacked100;
                            break;
                        default:
                            type = XL.XlChartType.xlLine;
                            break;
                    }
                    break;
                case LineChartType.LineMarkers:
                    switch (fillType)
                    {
                        case ChartFill.Stack:
                            type = XL.XlChartType.xlLineMarkersStacked;
                            break;
                        case ChartFill.Fill:
                            type = XL.XlChartType.xlLineMarkersStacked100;
                            break;
                        default:
                            type = XL.XlChartType.xlLineMarkers;
                            break;
                    }
                    break;
                case LineChartType.Area:
                    switch (fillType)
                    {
                        case ChartFill.Stack:
                            type = XL.XlChartType.xlAreaStacked;
                            break;
                        case ChartFill.Fill:
                            type = XL.XlChartType.xlAreaStacked100;
                            break;
                        default:
                            type = XL.XlChartType.xlArea;
                            break;
                    }
                    break;
                case LineChartType.Area3d:
                    switch (fillType)
                    {
                        case ChartFill.Stack:
                            type = XL.XlChartType.xl3DAreaStacked;
                            break;
                        case ChartFill.Fill:
                            type = XL.XlChartType.xl3DAreaStacked100;
                            break;
                        default:
                            type = XL.XlChartType.xl3DArea;
                            break;
                    }
                    break;
            }

            this.ComObj.Chart.ChartType = type;
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Chart | "+this.ComObj.Name;
        }

        #endregion

    }
}
