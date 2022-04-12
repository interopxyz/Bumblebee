using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace Bumblebee.Components.Graphics
{
    public class GH_Ex_CellBorder : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_CellBorder class.
        /// </summary>
        public GH_Ex_CellBorder()
          : base("Cell Border", "Cell Brd",
              "Border a cell",
              Constants.ShortName, Constants.SubGraphics)
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
            pManager.AddTextParameter("Cell Addresses", "A", "Cell addresses to modify. (ex. A1)", GH_ParamAccess.list);
            pManager.AddColourParameter("Colors", "C", "Border colors", GH_ParamAccess.list);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Weights", "W", "Border weight", GH_ParamAccess.list);
            pManager[3].Optional = true;
            pManager.AddIntegerParameter("Types", "T", "Border colors", GH_ParamAccess.list);
            pManager[4].Optional = true;
            pManager.AddIntegerParameter("Horizontal", "H", "Border colors", GH_ParamAccess.list);
            pManager[5].Optional = true;
            pManager.AddIntegerParameter("Vertical", "V", "Border colors", GH_ParamAccess.list);
            pManager[6].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[3];
            foreach (ExApp.BorderWeight value in Enum.GetValues(typeof(ExApp.BorderWeight)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramB = (Param_Integer)pManager[4];
            foreach (ExApp.LineType value in Enum.GetValues(typeof(ExApp.LineType)))
            {
                paramB.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramC = (Param_Integer)pManager[5];
            foreach (ExApp.HorizontalBorder value in Enum.GetValues(typeof(ExApp.HorizontalBorder)))
            {
                paramC.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramD = (Param_Integer)pManager[6];
            foreach (ExApp.VerticalBorder value in Enum.GetValues(typeof(ExApp.VerticalBorder)))
            {
                paramD.AddNamedValue(value.ToString(), (int)value);
            }
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
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;
            ExWorksheet worksheet = goo.ToWorksheet();

            List<string> addresses = new List<string>();
            DA.GetDataList(1, addresses);

            int countA = addresses.Count;

            List<Color> colors = new List<Color>();
            if(!DA.GetDataList(2, colors))colors.Add(Color.Black);
            int countB = colors.Count;

            for (int i = countB; i < countA; i++)
            {
                colors.Add(colors[countB - 1]);
            }

            List<int> weights = new List<int>();
            if (!DA.GetDataList(3, weights)) weights.Add(1);
            countB = weights.Count;

            for (int i = countB; i < countA; i++)
            {
                weights.Add(weights[countB - 1]);
            }

            List<int> types = new List<int>();
            if (!DA.GetDataList(4, types)) types.Add(1);

            for (int i = countB; i < countA; i++)
            {
                types.Add(types[countB - 1]);
            }

            List<int> horizontal = new List<int>();
            if (!DA.GetDataList(5, horizontal)) horizontal.Add(3);

            for (int i = countB; i < countA; i++)
            {
                horizontal.Add(horizontal[countB - 1]);
            }

            List<int> vertical = new List<int>();
            if (!DA.GetDataList(6, vertical)) vertical.Add(3);

            for (int i = countB; i < countA; i++)
            {
                vertical.Add(vertical[countB - 1]);
            }

            worksheet.Freeze();
            for (int i = 0; i < addresses.Count; i++)
            {
                worksheet.RangeBorder(addresses[i], addresses[i], colors[i], (ExApp.BorderWeight)weights[i], (ExApp.LineType)types[i], (ExApp.HorizontalBorder)horizontal[i], (ExApp.VerticalBorder)vertical[i]);
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
                return Properties.Resources.BB_Graphics_Border2_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("72ee99b6-6953-4621-a76d-0a802454cfc8"); }
        }
    }
}