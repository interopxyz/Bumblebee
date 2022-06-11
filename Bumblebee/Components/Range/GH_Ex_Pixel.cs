using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components.Range
{
    public class GH_Ex_Pixel : GH_Ex_Range_Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Pixel class.
        /// </summary>
        public GH_Ex_Pixel()
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
            pManager[2].Optional = true;
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
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;
            ExWorksheet worksheet = goo.ToWorksheet();

            string a = "A1";
            DA.GetData(1, ref a);

            string b = "A1";
            if (!DA.GetData(2, ref b)) b = a;

            Point3d ptA = worksheet.GetRangeMinPixel(a, b);
            Point3d ptB = worksheet.GetRangeMaxPixel(a, b);

            DA.SetData(0, worksheet);
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