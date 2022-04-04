using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components.Appearance
{
    public class GH_Ex_Vs_ColorCell : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Vs_ColorCell class.
        /// </summary>
        public GH_Ex_Vs_ColorCell()
          : base("Color Cell", "XL Cell Clr",
              "Color a cell",
              Constants.ShortName, Constants.SubGraphics)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Worksheet / Workbook / App", "Ws", "A Workbook, Worksheet, or Excel Application", GH_ParamAccess.item);
            pManager.AddTextParameter("Cell Addresses", "A", "Cell addresses to modify. (ex. A1)", GH_ParamAccess.list);
            pManager.AddColourParameter("Cell Colors", "C", "Cell colors", GH_ParamAccess.list);
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

            List<string> addresses = new List<string>();
            DA.GetDataList(1, addresses);

            List<Sd.Color> colors = new List<Sd.Color>();
            DA.GetDataList(2, colors);

            int countA = addresses.Count;
            int countC = colors.Count;
            for(int i = countC; i < countA; i++)
            {
                colors.Add(colors[countC - 1]);
            }

            worksheet.Freeze();
            for(int i =0;i<addresses.Count;i++)
            {
            worksheet.ColorCell(addresses[i], colors[i]);
            }
            worksheet.UnFreeze();
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
            get { return new Guid("8f76e2b5-87d7-49fa-8631-9f5d6bc31d29"); }
        }
    }
}