using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Dst_ReadData : GH_Ex_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Dt_ReadData class.
        /// </summary>
        public GH_Ex_Dst_ReadData()
          : base("Read Data", "XL Read",
              "Read data from Excel",
              Constants.ShortName, Constants.SubData)
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
            pManager[1].Optional = true;
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddTextParameter("Data", "D", "The data from the range where the Row is a branch with columns in a list.", GH_ParamAccess.tree);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool active = false;
            if (!DA.GetData(2, ref active)) return;

            if (active)
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

                GH_Structure<GH_String> data = range.ReadData();

                DA.SetData(0, range);
                DA.SetDataTree(1, data);

            }
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
                return Properties.Resources.BB_Sheet_Read_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("87c6d007-0d8b-4a71-bcde-8348ccb49563"); }
        }
    }
}