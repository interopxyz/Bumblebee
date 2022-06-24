using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components.Data
{
    public class GH_Ex_Dt_WriteFast : GH_Ex_Ws_Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_FastWrite class.
        /// </summary>
        public GH_Ex_Dt_WriteFast()
          : base("Fast Write Data", "XL Fast",
              "Fast Write data to excel",
              Constants.ShortName, Constants.SubData)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.secondary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager.AddGenericParameter("Cell", "C", "The cell to start writing to", GH_ParamAccess.item);
            pManager.AddTextParameter("Values", "V", "A datatree of values", GH_ParamAccess.tree);
            pManager.AddBooleanParameter("Activate", "A", "", GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddTextParameter("Start Address", "A", "The starting cell address of the range", GH_ParamAccess.item);
            pManager.AddTextParameter("Extent Address", "B", "The cell address at the extent of the range", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool active = false;
            if (!DA.GetData(3, ref active)) return;

            if (active)
            {
                IGH_Goo goo = null;
                if (!DA.GetData(0, ref goo)) return;
                ExWorksheet worksheet = goo.ToWorksheet();

                ExCell cell = new ExCell();
                DA.GetData(1, ref cell);

                string wbName = string.Empty;
                if(DA.GetData(1,ref wbName))
                {
                    workbook = app.GetWorkbook(wbName);
                }
                else
                {
                    workbook = app.GetActiveWorkbook();
                }

                string wsName = string.Empty;
                if (DA.GetData(2, ref wsName))
                {
                    worksheet = workbook.GetWorksheet(wsName);
                }
                else
                {
                    worksheet = workbook.GetActiveWorksheet();
                }

                GH_Structure<GH_String> ghData = new GH_Structure<GH_String> ();

                List<List<GH_String>> dataSet = new List<List<GH_String>>();
                if(!DA.GetDataTree(4, out ghData))return;

                foreach (List<GH_String> data in ghData.Branches)
                {
                    dataSet.Add(data);
                }

                string extent = worksheet.WriteData(dataSet, address);

                DA.SetData(0, app);
                DA.SetData(1, workbook);
                DA.SetData(2, worksheet);
                DA.SetData(3, address);
                DA.SetData(4, extent);
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
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("9e022976-510f-429f-a9d8-f659415168ef"); }
        }
    }
}