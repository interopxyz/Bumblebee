using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components
{
    public class GH_Ex_Rng_GetFont : GH_Ex_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Rng_GetFont class.
        /// </summary>
        public GH_Ex_Rng_GetFont()
          : base("Get Range Font", "Get Rng Font",
              "Get the Range font properties",
              Constants.ShortName, Constants.SubGraphics)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.tertiary; }
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
            pManager.AddTextParameter("Family Name", "F", "Font Family name", GH_ParamAccess.item);
            pManager.AddColourParameter("Color", "C", "Font color", GH_ParamAccess.item);
            pManager.AddNumberParameter("Size", "S", "Font size", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Justification", "J", "Text justifications", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Is Bold", "B", "Font Bold status", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Is Italic", "I", "Font Italic status", GH_ParamAccess.item);
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

            DA.SetData(0, range);
            DA.SetData(1, range.FontFamily);
            DA.SetData(2, range.FontColor);
            DA.SetData(3, range.FontSize);
            DA.SetData(4, (int)range.FontJustification);
            DA.SetData(5, range.Bold);
            DA.SetData(6, range.Italic);
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
                return Properties.Resources.BB_Graphics_GetFont_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("a03bb7b7-727b-4ae6-86a9-359637c85b5d"); }
        }
    }
}