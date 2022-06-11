using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components.Appearance
{
    public class GH_Ex_CellColor : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Vs_ColorCell class.
        /// </summary>
        public GH_Ex_CellColor()
          : base("Cell Color", "Cell Clr",
              "Color a cell",
              Constants.ShortName, Constants.SubGraphics)
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
            pManager.AddGenericParameter("Worksheet / Workbook / App", "Ws", "A Workbook, Worksheet, or Excel Application", GH_ParamAccess.item);
            pManager.AddTextParameter("Cell Addresses", "A", "Cell addresses to modify. (ex. A1)", GH_ParamAccess.list);
            pManager.AddColourParameter("Cell Colors", "C", "Cell colors", GH_ParamAccess.list);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Worksheet", "Ws", "The updated worksheet", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool isSingle = true;
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;
            ExWorksheet worksheet = goo.ToWorksheet();

            List<string> addresses = new List<string>();
            DA.GetDataList(1, addresses);

            List<Sd.Color> colors = new List<Sd.Color>();
            DA.GetDataList(2, colors);

            int countA = addresses.Count;
            int countB = colors.Count;
            if (countB > 1) isSingle = false;
            for (int i = countB; i < countA; i++)
            {
                colors.Add(colors[countB - 1]);
            }

            worksheet.Freeze();
            if (isSingle)
            {
                worksheet.RangeColor(addresses[0], addresses[countA-1], colors[0]);
            }
            else
            {
                for (int i = 0; i < addresses.Count; i++)
                {
                    worksheet.RangeColor(addresses[i], addresses[i], colors[i]);
                }
            }
            worksheet.UnFreeze();

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
                return Properties.Resources.BB_Graphics_Fill3_01;
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