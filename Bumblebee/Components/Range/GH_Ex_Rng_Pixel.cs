using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Rng_Pixel : GH_Ex_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Pixel class.
        /// </summary>
        public GH_Ex_Rng_Pixel()
          : base("Pixel Range", "Rng Pxl",
              "Gets the minimum and maximum pixel location",
              Constants.ShortName, Constants.SubRange)
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
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddPointParameter("Point A", "A", "The pixel location of the ranges upper left hand corner", GH_ParamAccess.item);
            pManager.AddPointParameter("Point B", "B", "The pixel location of the ranges lower right hand corner", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooS = null;
            if (!DA.GetData(0, ref gooS)) return;
            ExWorksheet worksheet = null;
            if (!gooS.TryGetWorksheet(ref worksheet)) return;

            IGH_Goo gooR = null;
            DA.GetData(1, ref gooR);
            ExRange range = null;
            if (!gooR.TryGetRange(ref range, worksheet)) return;

            Point3d ptA = range.GetMinPixel();
            Point3d ptB = range.GetMaxPixel();

            DA.SetData(0, range);
            DA.SetData(1, ptA);
            DA.SetData(2, ptB);
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
                return Properties.Resources.BB_Range_Pixel2_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("5969629c-32c8-441a-a76d-6695dc59027a"); }
        }
    }
}