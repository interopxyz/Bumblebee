using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Wks_Get : GH_Ex_Wbk__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Ws_Get class.
        /// </summary>
        public GH_Ex_Wks_Get()
          : base("Get Worksheet", "Get Worksheet",
              "Get a worksheet by name or the active worksheet from a workbook or application",
              Constants.ShortName, Constants.SubWorkSheets)
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
            pManager[0].Optional = true;
            pManager.AddGenericParameter("Name", "N", "The name of an active Worksheet", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Worksheet.Name, Constants.Worksheet.NickName, Constants.Worksheet.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooB = null;
            DA.GetData(0, ref gooB);
            ExWorkbook workbook = null;
            gooB.TryGetWorkbook(ref workbook);

            IGH_Goo gooS = null;
            DA.GetData(1, ref gooS);
            ExWorksheet worksheet = null;
            if (!gooS.TryGetWorksheet(ref worksheet, workbook)) return;

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
                return Properties.Resources.BB_Sheet_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("d8b5f26b-310a-4bdd-9f9e-f7c01d7faf43"); }
        }
    }
}