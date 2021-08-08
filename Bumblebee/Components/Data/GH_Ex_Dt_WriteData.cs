using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Dt_WriteData : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WriteData class.
        /// </summary>
        public GH_Ex_Dt_WriteData()
          : base("Write Data", "XL Write",
              "Write data to excel",
              Constants.ShortName, "Data")
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.secondary
                    ; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Worksheet / Workbook / App", "Ws", "A Workbook, Worksheet, or Excel Application", GH_ParamAccess.item);
            pManager.AddTextParameter("Cell Address", "A", "The cell address to start writing to in standard address format. (ex. A1", GH_ParamAccess.item, "A1");
            pManager.AddGenericParameter("Data", "Da", "The dataset to write to excel", GH_ParamAccess.list);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
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

            if (goo.CastTo<ExWorksheet>(out worksheet))
            {
            }
            else if (goo.CastTo<ExWorkbook>(out workbook))
            {
                worksheet = workbook.GetActiveWorksheet();
            }
            else if (goo.CastTo<ExApp>(out app))
            {
                worksheet = app.GetActiveWorksheet();
            }

            string address = "A1";
            DA.GetData(1, ref address);

            List<ExData> data = new List<ExData>();
            if (!DA.GetDataList(2, data)) return;

            worksheet.WriteData(data, address);

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
            get { return new Guid("b372e027-76eb-4ec8-a49c-76fcb2f7985b"); }
        }
    }
}