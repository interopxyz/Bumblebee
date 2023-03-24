﻿using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components
{
    public class GH_Ex_Cht_ChartScatter : GH_Ex_Cht__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Ch_ChartScatter class.
        /// </summary>
        public GH_Ex_Cht_ChartScatter()
          : base("Scatter Chart", "Scatter Chart",
              "Add a Scatter Chart object",
              Constants.ShortName, Constants.SubChart)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.primary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager.AddIntegerParameter("Chart Type", "T", "The chart type", GH_ParamAccess.item, 0);
            pManager[5].Optional = true;
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item, false);
            pManager[6].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[5];
            foreach (ScatterChartType value in Enum.GetValues(typeof(ScatterChartType)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooS = null;
            DA.GetData(0, ref gooS);
            ExWorksheet worksheet = new ExWorksheet();
            bool hasWs = gooS.TryGetWorksheet(ref worksheet);

            IGH_Goo gooR = null;
            if (!DA.GetData(1, ref gooR)) return;
            ExRange range = new ExRange();
            if (!gooR.TryGetRange(ref range, worksheet)) return;
            if (!hasWs) worksheet = range.Worksheet;

            string name = GetInstanceName();
            DA.GetData(2, ref name);

            Rectangle3d boundary = new Rectangle3d(Plane.WorldXY, new Point3d(250, 250, 0), new Point3d(500, 500, 0));
            DA.GetData(3, ref boundary);

            bool flip = false;
            DA.GetData(4, ref flip);

            int type = 0;
            DA.GetData(5, ref type);

            bool activate = false;
            DA.GetData(6, ref activate);

            if (activate)
            {
                ExChart chart = new ExChart(name, range, flip, boundary);
                chart.SetScatterChart((ScatterChartType)type);
                DA.SetData(1, chart);
            }

            DA.SetData(0, range);
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
                return Properties.Resources.BB_Chart_Scatter_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("6c5e65ef-3c69-4daf-8ead-d6796aa0dc43"); }
        }
    }
}