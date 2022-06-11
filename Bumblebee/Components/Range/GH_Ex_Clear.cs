using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components.Range
{
    public class GH_Ex_Clear : GH_Ex_Range_Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Clear class.
        /// </summary>
        public GH_Ex_Clear()
          : base("Clear Range", "Rng Pxl",
              "Clears the contents or formatting of a range",
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
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Content", "C", "If true, the range content will be cleared", GH_ParamAccess.item, true);
            pManager[3].Optional = true;
            pManager.AddBooleanParameter("Formatting", "F", "If true, the range format will be cleared", GH_ParamAccess.item, false);
            pManager[4].Optional = true;
            pManager.AddBooleanParameter("Activate", "_A", "If true, the component will be activated", GH_ParamAccess.item, false);
            pManager[5].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddTextParameter("Start Address", "A", "The starting cell address of the range", GH_ParamAccess.item);
            pManager.AddTextParameter("Extent Address", "B", "The cell address that sets the bounding extent of the range", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Status", "S", "Returns the status of the activate input", GH_ParamAccess.item);
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
            if(!DA.GetData(1, ref a)) a = worksheet.GetFirstUsedCell();

            string b = "A1";
            if (!DA.GetData(2, ref b)) b = worksheet.GetLastUsedCell();

            bool content = true;
            DA.GetData(3, ref content);

            bool format = false;
            DA.GetData(4, ref format);

            bool activate = false;
            DA.GetData(5, ref activate);

            if (activate)
            {
                if (content) worksheet.ClearContent(a, b);
                if (format) worksheet.ClearFormat(a, b);
            }

            DA.SetData(0, worksheet);
            DA.SetData(1, a);
            DA.SetData(2, b);
            DA.SetData(3, activate);
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
                return Properties.Resources.BB_Range_Clear2_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("13b12cdd-bc51-4df6-bace-e190a894474b"); }
        }
    }
}