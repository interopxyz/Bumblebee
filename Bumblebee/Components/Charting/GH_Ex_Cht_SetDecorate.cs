using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components
{
    public class GH_Ex_Cht_SetDecorate : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Ch_Decorate class.
        /// </summary>
        public GH_Ex_Cht_SetDecorate()
          : base("Chart Decorators", "Decorate Chart",
              "Update Chart Decorations",
              Constants.ShortName, Constants.SubChart)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.secondary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Chart.Name, Constants.Chart.NickName, Constants.Chart.Input, GH_ParamAccess.item);
            pManager.AddTextParameter("Title", "T", "", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddIntegerParameter("Legend Location", "L", "The location of the legend", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Data Label", "D", "The data label type", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddIntegerParameter("Grid X", "Gx", "The X axis Grid settings", GH_ParamAccess.item);
            pManager[4].Optional = true;
            pManager.AddIntegerParameter("Grid Y", "Gy", "The Y axis Grid settings", GH_ParamAccess.item);
            pManager[5].Optional = true;
            pManager.AddTextParameter("Axis X", "Ax", "An optional X axis label", GH_ParamAccess.item);
            pManager[6].Optional = true;
            pManager.AddTextParameter("Axis Y", "Ay", "An optional Y axis label", GH_ParamAccess.item);
            pManager[7].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[2];
            foreach (LegendLocations value in Enum.GetValues(typeof(LegendLocations)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramB = (Param_Integer)pManager[3];
            foreach (LabelType value in Enum.GetValues(typeof(LabelType)))
            {
                paramB.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramC = (Param_Integer)pManager[4];
            foreach (GridType value in Enum.GetValues(typeof(GridType)))
            {
                paramC.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramD = (Param_Integer)pManager[5];
            foreach (GridType value in Enum.GetValues(typeof(GridType)))
            {
                paramD.AddNamedValue(value.ToString(), (int)value);
            }

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Chart.Name, Constants.Chart.NickName, Constants.Chart.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            ExChart chart = null;
            if (!DA.GetData(0, ref chart)) return;
            chart = new ExChart(chart);

            string title = "";
            if (DA.GetData(1, ref title)) chart.SetTitle(title);

            int legend = 0;
            if (DA.GetData(2, ref legend)) chart.SetLegend((LegendLocations)legend);

            int dataLabel = 0;
            if (DA.GetData(3, ref dataLabel)) chart.SetLabel((LabelType)dataLabel);

            int gridX = 0;
            if (DA.GetData(4, ref gridX)) chart.SetGridX( (GridType)gridX);

            int gridY = 0;
            if (DA.GetData(5, ref gridY)) chart.SetGridY((GridType)gridY);

            string axisX = "";
            if (DA.GetData(6, ref axisX)) chart.SetAxisX(axisX);

            string axisY = "";
            if (DA.GetData(7, ref axisY)) chart.SetAxisY(axisY);

            DA.SetData(0, chart);
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return Properties.Resources.BB_Chart_Decorate_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("31988a4f-b6f4-46f6-93c7-ce2c28d05ca8"); }
        }
    }
}