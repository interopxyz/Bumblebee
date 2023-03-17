using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Rng_Resize : GH_Ex_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Resize class.
        /// </summary>
        public GH_Ex_Rng_Resize()
          : base("Resize Cell", "Cell Size",
              "Set the Column width and Row height for a Range",
              Constants.ShortName, Constants.SubRange)
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
            pManager[1].Optional = true;
            pManager.AddIntegerParameter("Column Width", "C", "The column width", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Row Height", "R", "The column width", GH_ParamAccess.item);
            pManager[3].Optional = true;
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
            DA.GetData(1, ref gooR);
            ExRange range = null;
            if (!gooR.TryGetRange(ref range, worksheet)) return;
            if (!hasWs) worksheet = range.Worksheet;

            //worksheet.Freeze();

            int width = 10;
            if(DA.GetData(2, ref width)) range.Width = width;

            int height = 10;
            if(DA.GetData(3, ref height)) range.Height=height;

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
                return Properties.Resources.BB_Cell_SizeD_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("3220ae58-927b-4416-8306-b96fe13a2bd2"); }
        }
    }
}