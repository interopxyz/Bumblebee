using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components.Range
{
    public class GH_Ex_Extents : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Extents class.
        /// </summary>
        public GH_Ex_Extents()
          : base("Range Extents", "Rng Bnd",
              "Gets the minimum and maximum active Cell address for the sheet",
              Constants.ShortName, Constants.SubRange)
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
            pManager.AddGenericParameter("Worksheet / Workbook / App", "Ws", "A Workbook, Worksheet, or Excel Application", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Worksheet", "Ws", "The updated worksheet", GH_ParamAccess.item);
            pManager.AddTextParameter("Start Address", "A", "The starting cell address of the range", GH_ParamAccess.item);
            pManager.AddTextParameter("Extent Address", "B", "The cell address at the extent of the range", GH_ParamAccess.item);
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

            string start = worksheet.GetFirstUsedCell();
            string extent = worksheet.GetLastUsedCell();

            DA.SetData(0, worksheet);
            DA.SetData(1, start);
            DA.SetData(2, extent);
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
                return Properties.Resources.BB_Extents_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("b588c26a-e6f3-437a-8392-5b5611987ca1"); }
        }
    }
}