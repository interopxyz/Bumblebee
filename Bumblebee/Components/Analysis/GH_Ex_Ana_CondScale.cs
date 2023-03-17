using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components
{
    public class GH_Ex_Ana_CondScale : GH_Ex_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_An_CondScale class.
        /// </summary>
        public GH_Ex_Ana_CondScale()
          : base("Conditional Scale", "Scale",
              "Add conditional formatting colors to a Range based on relative values",
              Constants.ShortName, Constants.SubAnalysis)
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
            base.RegisterInputParams(pManager);
            pManager[1].Optional = true;
            pManager.AddNumberParameter("Parameter", "P", "The parameter of the midpoint of a 3 color gradient", GH_ParamAccess.item, 0.5);
            pManager[2].Optional = true;
            pManager.AddColourParameter("Gradient Color 1", "C0", "The first color of the gradient", GH_ParamAccess.item, Sd.Color.LightGray);
            pManager[3].Optional = true;
            pManager.AddColourParameter("Gradient Color 2", "C1", "The second color of the gradient", GH_ParamAccess.item, Sd.Color.Gray);
            pManager[4].Optional = true;
            pManager.AddColourParameter("Gradient Color 3", "C2", "The third color of the gradient", GH_ParamAccess.item);
            pManager[5].Optional = true;
            pManager.AddBooleanParameter("Clear", "_X", "If true, the existing conditions will be cleared", GH_ParamAccess.item, false);
            pManager[6].Optional = true;
            pManager.AddBooleanParameter("Activate", "_A", "If true, the condition will be applied", GH_ParamAccess.item, false);
            pManager[7].Optional = true;
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
            DA.GetData(1, ref gooR);
            ExRange range = new ExRange();
            if (!gooR.TryGetRange(ref range, worksheet)) return;
            if (!hasWs) worksheet = range.Worksheet;

            double mid = 0.5;
            DA.GetData(2, ref mid);

            Sd.Color a = Sd.Color.LightGray;
            DA.GetData(3, ref a);

            Sd.Color b = Sd.Color.Gray;
            DA.GetData(4, ref b);

            Sd.Color c = Sd.Color.DarkGray;
            bool isThree = DA.GetData(5, ref c);

            bool clear = false;
            DA.GetData(6, ref clear);

            bool activate = false;
            DA.GetData(7, ref activate);

            if (activate)
            {
                if (clear) range.ClearConditions();
                if (isThree)
                {
                    range.AddConditionalScale(a,b,c,mid);
                }
                else
                {
                    range.AddConditionalScale(a, b);
                }
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
                return Properties.Resources.BB_Cond_Scale_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("7c251907-0e3b-4c57-b948-98ea7c76b68f"); }
        }
    }
}