using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Rng_Clear : GH_Ex_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Clear class.
        /// </summary>
        public GH_Ex_Rng_Clear()
          : base("Clear Range", "Rng Pxl",
              "Clears the contents or formatting of a Range",
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
            pManager.AddBooleanParameter("Content", "C", "If true, the range content will be cleared", GH_ParamAccess.item, true);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Formatting", "F", "If true, the range format will be cleared", GH_ParamAccess.item, false);
            pManager[3].Optional = true;
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item, false);
            pManager[4].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddBooleanParameter("Status", "S", "Returns the status of the activate input", GH_ParamAccess.item);
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
            DA.GetData(1, ref gooR);
            ExRange range = new ExRange();
            if (!gooR.TryGetRange(ref range, worksheet)) return;
            if (!hasWs) worksheet = range.Worksheet;

            bool content = true;
            DA.GetData(2, ref content);

            bool format = false;
            DA.GetData(3, ref format);

            bool activate = false;
            DA.GetData(4, ref activate);

            if (activate)
            {
                if (content) range.ClearContent();
                if (format) range.ClearFormat();
            }

            DA.SetData(0, range);
            DA.SetData(1, activate);
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