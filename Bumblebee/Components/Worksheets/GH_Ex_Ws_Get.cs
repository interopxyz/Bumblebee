using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Ws_Get : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Ws_Get class.
        /// </summary>
        public GH_Ex_Ws_Get()
          : base("Get Worksheet", "Get Worksheet",
              "Get a worksheet by name or the active worksheet from a workbook or application",
              Constants.ShortName, "Worksheet")
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
            pManager.AddGenericParameter("Workbook / XlApp", "Wb", "A Workbook or Excel Application", GH_ParamAccess.item);
            pManager.AddTextParameter("Name", "N", "The name of an active Workbook", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Worksheet", "Ws", "The Excel Worksheet object", GH_ParamAccess.item);
            pManager.AddGenericParameter("Workbook", "Wb", "The Excel Workbook object", GH_ParamAccess.item);
            pManager.AddGenericParameter("App", "App", "The parent application.", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            ExApp app = null;
            ExWorkbook workbook = null;

            ExWorksheet worksheet = null;

            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;

            string name = string.Empty;
            if (DA.GetData(1, ref name))
            {
                if (goo.CastTo<ExApp>(out app))
                {
                    worksheet = app.GetWorksheet(name);
                }
                else if (goo.CastTo<ExWorkbook>(out workbook))
                {
                    worksheet = workbook.GetWorksheet(name);
                }
            }
            else
            {
                if (goo.CastTo<ExApp>(out app))
                {
                    worksheet = app.GetActiveWorksheet();
                }
                else if (goo.CastTo<ExWorkbook>(out workbook))
                {
                    worksheet = workbook.GetActiveWorksheet();
                }
            }
            workbook = worksheet.Workbook;
            app = workbook.ParentApp;

            DA.SetData(0, worksheet);
            DA.SetData(1, workbook);
            DA.SetData(2, app);
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