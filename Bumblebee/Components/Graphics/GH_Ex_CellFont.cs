using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace Bumblebee.Components.Graphics
{
    public class GH_Ex_CellFont : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_CellFont class.
        /// </summary>
        public GH_Ex_CellFont()
          : base("Cell Font", "Cell Font",
              "Change a cell font",
              Constants.ShortName, Constants.SubGraphics)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.tertiary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Worksheet / Workbook / App", "Ws", "A Workbook, Worksheet, or Excel Application", GH_ParamAccess.item);
            pManager.AddTextParameter("Cell Addresses", "A", "Cell addresses to modify. (ex. A1)", GH_ParamAccess.list);
            pManager.AddTextParameter("Font Names", "F", "Cell font names", GH_ParamAccess.list);
            pManager[2].Optional = true;
            pManager.AddColourParameter("Font Colors", "C", "Cell font colors", GH_ParamAccess.list);
            pManager[3].Optional = true;
            pManager.AddNumberParameter("Font Sizes", "S", "Cell font numbers", GH_ParamAccess.list);
            pManager[4].Optional = true;
            pManager.AddIntegerParameter("Font Justifications", "J", "Cell font justifications", GH_ParamAccess.list);
            pManager[5].Optional = true;
            pManager.AddBooleanParameter("Is Bold", "B", "Cell font Bold status", GH_ParamAccess.list);
            pManager[6].Optional = true;
            pManager.AddBooleanParameter("Is Italic", "I", "Cell font Italic status", GH_ParamAccess.list);
            pManager[7].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[5];
            foreach (ExApp.Justification value in Enum.GetValues(typeof(ExApp.Justification)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
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
            bool isSingle = true;
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;
            ExWorksheet worksheet = goo.ToWorksheet();

            List<string> addresses = new List<string>();
            DA.GetDataList(1, addresses);

            int countA = addresses.Count;

            List<string> names = new List<string>();
            if (!DA.GetDataList(2, names)) names.Add("Arial");
            int countB = names.Count;
            if (countB > 1) isSingle = false;
            for (int i = countB; i < countA; i++)
            {
                names.Add(names[countB - 1]);
            }

            List<Color> colors = new List<Color>();
            if (!DA.GetDataList(3, colors)) colors.Add(Color.Black);
            countB = colors.Count;
            if (countB > 1) isSingle = false;
            for (int i = countB; i < countA; i++)
            {
                colors.Add(colors[countB - 1]);
            }

            List<double> sizes = new List<double>();
            if (!DA.GetDataList(4, sizes)) sizes.Add(10.0);
            countB = sizes.Count;
            if (countB > 1) isSingle = false;
            for (int i = countB; i < countA; i++)
            {
                sizes.Add(sizes[countB - 1]);
            }

            List<int> justifications = new List<int>();
            if (!DA.GetDataList(5, justifications)) justifications.Add(1);
            countB = justifications.Count;
            if (countB > 1) isSingle = false;
            for (int i = countB; i < countA; i++)
            {
                justifications.Add(justifications[countB - 1]);
            }

            List<bool> isBold = new List<bool>();
            if (!DA.GetDataList(6, isBold)) isBold.Add(false);
            countB = isBold.Count;
            if (countB > 1) isSingle = false;
            for (int i = countB; i < countA; i++)
            {
                isBold.Add(isBold[countB - 1]);
            }

            List<bool> isItalic = new List<bool>();
            if (!DA.GetDataList(7, isItalic)) isItalic.Add(false);
            countB = isItalic.Count;
            if (countB > 1) isSingle = false;
            for (int i = countB; i < countA; i++)
            {
                isItalic.Add(isItalic[countB - 1]);
            }

            worksheet.Freeze();
            if (isSingle)
            {
                worksheet.RangeFont(addresses[0], addresses[countA-1], names[0], sizes[0], colors[0], (ExApp.Justification)justifications[0], isBold[0], isItalic[0]);
            }
            else
            {
                for (int i = 0; i < addresses.Count; i++)
                {
                    worksheet.RangeFont(addresses[i], addresses[i], names[i], sizes[i], colors[i], (ExApp.Justification)justifications[i], isBold[i], isItalic[i]);
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
                return Properties.Resources.BB_Graphics_Font_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("03406742-5ebf-41e0-ab46-6e15f8aa5766"); }
        }
    }
}