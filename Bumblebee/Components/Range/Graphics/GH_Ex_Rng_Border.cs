using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components
{
    public class GH_Ex_Rng_Border : GH_Ex_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_CellBorder class.
        /// </summary>
        public GH_Ex_Rng_Border()
          : base("Range Border", "Rng Brd",
              "Sets the Range border properties",
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
            pManager.AddColourParameter("Color", "C", "Border color", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Weight", "W", "Border weight", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddIntegerParameter("Type", "T", "Border type", GH_ParamAccess.item);
            pManager[4].Optional = true;
            pManager.AddIntegerParameter("Horizontal", "H", "Border horizontal side", GH_ParamAccess.item);
            pManager[5].Optional = true;
            pManager.AddIntegerParameter("Vertical", "V", "Border vertical side", GH_ParamAccess.item);
            pManager[6].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[3];
            foreach (BorderWeight value in Enum.GetValues(typeof(BorderWeight)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramB = (Param_Integer)pManager[4];
            foreach (LineType value in Enum.GetValues(typeof(LineType)))
            {
                paramB.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramC = (Param_Integer)pManager[5];
            foreach (HorizontalBorder value in Enum.GetValues(typeof(HorizontalBorder)))
            {
                paramC.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramD = (Param_Integer)pManager[6];
            foreach (VerticalBorder value in Enum.GetValues(typeof(VerticalBorder)))
            {
                paramD.AddNamedValue(value.ToString(), (int)value);
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

            Sd.Color color = Sd.Color.Black;
            DA.GetData(2, ref color);

            int weight = 1;
            DA.GetData(3, ref weight);

            int type = 1;
            DA.GetData(4,ref type);

            int horizontal = 3;
            DA.GetData(5, ref horizontal);

            int vertical = 3;
            DA.GetData(6, ref vertical);

            range.SetBorder(color, (BorderWeight)weight, (LineType)type, (HorizontalBorder)horizontal, (VerticalBorder)vertical);

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
                return Properties.Resources.BB_Graphics_Border;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("72ee99b6-6953-4621-a76d-0a802454cfc8"); }
        }
    }
}