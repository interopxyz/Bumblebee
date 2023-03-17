using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components
{
    public class GH_Ex_Ana_SparkLine : GH_Ex_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the Ex_Gh_An_Sparkline class.
        /// </summary>
        public GH_Ex_Ana_SparkLine()
          : base("Sparkline", "Spark",
              "Add a Sparkline",
              Constants.ShortName, Constants.SubAnalysis)
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
            pManager.AddGenericParameter("Placement", "P", "A single Cell Range to place the sparkline", GH_ParamAccess.item);
            pManager.AddColourParameter("Color", "C", "The Sparkline color", GH_ParamAccess.item, Sd.Color.Black);
            pManager[3].Optional = true;
            pManager.AddNumberParameter("Weight", "W", "the Sparkline weight", GH_ParamAccess.item, 1);
            pManager[4].Optional = true;
            pManager.AddBooleanParameter("Activate", "_A", "If true, the component will be activated", GH_ParamAccess.item, false);
            pManager[5].Optional = true;
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

            IGH_Goo gooP = null;
            if (!DA.GetData(2, ref gooP)) return;
            ExRange placement = new ExRange();
            if (!gooR.TryGetRange(ref placement, worksheet)) return;

            Sd.Color color = Sd.Color.Black;
            DA.GetData(3, ref color);

            double weight = 1.0;
            DA.GetData(4, ref weight);

            bool activate = false;
            DA.GetData(5, ref activate);

            if (activate) range.AddSparkLine(placement, color, weight);

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
                return Properties.Resources.BB_Cell_Sparkline_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("24529e60-47b1-4201-94c2-58f14682bab6"); }
        }
    }
}