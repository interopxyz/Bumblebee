using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Ws_GetAll : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Ws_GetAll class.
        /// </summary>
        public GH_Ex_Ws_GetAll()
          : base("Get Worksheets", "Get Worksheets",
              "Gets all the worksheets in a workbook or application",
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
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Worksheet", "Ws", "The Excel Worksheet object", GH_ParamAccess.list);
            pManager.AddGenericParameter("Workbook", "Wb", "The Excel Workbook object", GH_ParamAccess.list);
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

            List<ExWorksheet> worksheets = new List<ExWorksheet>();

            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;

            if (goo.CastTo<ExApp>(out app))
            {
                worksheets = app.GetWorksheets();
            }
            else if (goo.CastTo<ExWorkbook>(out workbook))
            { 
                worksheets = workbook.GetWorksheets();
            }

            List<ExWorkbook> workbooks = new List<ExWorkbook>();
            foreach(ExWorksheet sheet in worksheets)
            {
                workbooks.Add(sheet.Workbook);
            }

            if (workbooks.Count > 0) app = workbooks[0].ParentApp;

            DA.SetDataList(0, worksheets);
            DA.SetDataList(1, workbooks);
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
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("fa6fb9e1-1e5f-42f6-b6a9-a66ef15141da"); }
        }
    }
}