using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components
{
    public class GH_Ex_Rng_GetBorder : GH_Ex_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Rng_GetBorder class.
        /// </summary>
        public GH_Ex_Rng_GetBorder()
          : base("Get Range Border", "Get Rng Brd",
              "Get the Range border properties ",
              Constants.ShortName, Constants.SubGraphics)
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
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddColourParameter("Color", "C", "Border color", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Weight", "W", "Border weight", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Type", "T", "Border type", GH_ParamAccess.item);
            //pManager.AddIntegerParameter("Horizontal", "H", "Border horizontal side", GH_ParamAccess.item);
            //pManager.AddIntegerParameter("Vertical", "V", "Border vertical side", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooS = null;
            DA.GetData(0, ref gooS);
            ExWorksheet worksheet = null;
            bool hasWs = gooS.TryGetWorksheet(ref worksheet);

            IGH_Goo gooR = null;
            if (!DA.GetData(1, ref gooR)) return;
            ExRange range = null;
            if (!gooR.TryGetRange(ref range, worksheet)) return;
            if (!hasWs) worksheet = range.Worksheet;

            DA.SetData(0, range);
            DA.SetData(1, range.BorderColor);
            DA.SetData(2, (int)range.Weight);
            DA.SetData(3, (int)range.LineType);
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
                return Properties.Resources.BB_Graphics_GetBorder_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("17be318b-4d39-4c0b-af37-7357ef39ec39"); }
        }
    }
}