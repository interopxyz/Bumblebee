using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Resize : GH_Ex_Range_Base
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
            pManager[3].Optional = true;
            pManager.AddIntegerParameter("Row Height", "R", "The column width", GH_ParamAccess.item);
            pManager[4].Optional = true;
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

            string a = "A1";
            if (!DA.GetData(1, ref a)) return;

            string b = "A1";
            if (!DA.GetData(2, ref b)) b = a;

            int width = 10;
            DA.GetData(3, ref width);

            int height = 15;
            DA.GetData(4, ref height);

            worksheet.Freeze();
            worksheet.ResizeRangeCells(a, b, width, height);
            worksheet.UnFreeze();

            DA.SetData(0, worksheet);
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
                return Properties.Resources.BB_Cell_Size_01;
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