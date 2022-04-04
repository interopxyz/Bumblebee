using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Wb_Get : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Workbook class.
        /// </summary>
        public GH_Ex_Wb_Get()
          : base("Get Workbook", "Get Workbook",
              "Get a workbook by name or active workbook",
              Constants.ShortName, Constants.SubWorkBooks)
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
            pManager.AddGenericParameter("App", "App", "The Excel application", GH_ParamAccess.item);
            pManager.AddTextParameter("Name", "N", "The name of an active Workbook", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
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
            string name = string.Empty;

            ExWorkbook workbook = null;

            if (!DA.GetData<ExApp>(0, ref app)) return;
            if (DA.GetData(1, ref name))
            {
                workbook = app.GetWorkbook(name);
                if (workbook == null) { }
            }
            else
            {
                workbook = app.GetActiveWorkbook();
            }

            DA.SetData(0, workbook);
            DA.SetData(1, app);
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
                return Properties.Resources.BB_Book_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("a865970b-6e4f-43a1-8004-308c7b7db5a1"); }
        }
    }
}