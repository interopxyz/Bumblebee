using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components
{
    public class GH_Ex_Rng_SetFont : GH_Ex_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_CellFont class.
        /// </summary>
        public GH_Ex_Rng_SetFont()
          : base("Range Font", "Rng Font",
              "Sets the Range font properties",
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
            pManager.AddTextParameter("Family Name", "F", "Font Family name", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddColourParameter("Color", "C", "Font color", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddNumberParameter("Size", "S", "Font size", GH_ParamAccess.item);
            pManager[4].Optional = true;
            pManager.AddIntegerParameter("Justification", "J", "Text justifications", GH_ParamAccess.item);
            pManager[5].Optional = true;
            pManager.AddBooleanParameter("Is Bold", "B", "Font Bold status", GH_ParamAccess.item);
            pManager[6].Optional = true;
            pManager.AddBooleanParameter("Is Italic", "I", "Font Italic status", GH_ParamAccess.item);
            pManager[7].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[5];
            foreach (Justification value in Enum.GetValues(typeof(Justification)))
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
            ExWorksheet worksheet = null;
            bool hasWs = gooS.TryGetWorksheet(ref worksheet);

            IGH_Goo gooR = null;
            if (!DA.GetData(1, ref gooR)) return;
            ExRange range = null;
            if (!gooR.TryGetRange(ref range, worksheet)) return;
            if (!hasWs) worksheet = range.Worksheet;

            //worksheet.Freeze();

            string family = "Arial";
            if (DA.GetData(2,ref family)) range.FontFamily = family;

            Sd.Color color = Sd.Color.Black;
            if (DA.GetData(3, ref color)) range.FontColor = color;

            double size = 10.0;
            if (DA.GetData(4, ref size)) range.FontSize = size;

            int justifications = 1;
            if (DA.GetData(5, ref justifications)) range.FontJustification = (Justification)justifications;

            bool isBold = false;
            if (DA.GetData(6,ref isBold)) range.Bold = isBold;

            bool isItalic = false;
            if (DA.GetData(7, ref isItalic)) range.Italic = isItalic;

            //worksheet.UnFreeze();

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
                return Properties.Resources.BB_Graphics_Font_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("03406742-5ebf-41e0-ab46-6e15f8aa5766"); }
        }
    }
}