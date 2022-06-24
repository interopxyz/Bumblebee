using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Resize : GH_Ex_Rng_Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Resize class.
        /// </summary>
        public GH_Ex_Resize()
          : base("Resize Cell", "Cell Size",
              "Set the Column width and row height for a range",
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
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;
            ExWorksheet worksheet = goo.ToWorksheet();

            IGH_Goo gooR = null;
            if (!DA.GetData(1, ref gooR)) return;
            ExRange range = new ExRange();
                if(!gooR.TryGetRange(ref range))return;

            worksheet.Freeze();

            int width = 10;
            if(DA.GetData(2, ref width)) worksheet.RangeWidth(range, width); ;

            int height = 10;
            if(DA.GetData(3, ref height)) worksheet.RangeHeight(range, height); ;

            worksheet.UnFreeze();

            DA.SetData(0, worksheet);
            DA.SetData(1, range);
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