using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Cel_AddressesRange : GH_Ex_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_AddressesRange class.
        /// </summary>
        public GH_Ex_Cel_AddressesRange()
          : base("Range Addresses", "Addresses",
              "Generate a list of Cells from a Range.",
              Constants.ShortName, Constants.SubCell)
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
            pManager.AddBooleanParameter("Flip", "F", "If true, cells are listed by column. If false, by row", GH_ParamAccess.item, true);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Outputs, GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooS = null;
            DA.GetData(0, ref gooS);
            ExWorksheet worksheet = new ExWorksheet();
            bool hasWs = gooS.TryGetWorksheet(ref worksheet);

            IGH_Goo gooR = null;
            if (!DA.GetData(1, ref gooR)) return;
            ExRange range = new ExRange();
            if (!gooR.TryGetRange(ref range, worksheet)) return;
            if (!hasWs) worksheet = range.Worksheet;

            ExCell a = range.Start;
            ExCell b = range.Extent;

            bool flip = true;
            DA.GetData(2, ref flip);

            int xa = Math.Min(a.Column, b.Column);
            int xb = Math.Max(a.Column, b.Column) + 1;

            int ya = Math.Min(a.Row, b.Row);
            int yb = Math.Max(a.Row, b.Row) + 1;

            List<ExCell> cells = new List<ExCell>();

            if (flip)
            {
                for (int i = xa; i < xb; i++)
                {
                    for (int j = ya; j < yb; j++)
                    {
                        cells.Add(new ExCell(i, j));
                    }
                }
            }
            else
            {
                for (int j = xa; j < xb; j++)
                {
                    for (int i = ya; i < yb; i++)
                    {
                        cells.Add(new ExCell(i, j));
                    }
                }
            }

            DA.SetData(0, range);
            DA.SetDataList(1, cells);
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
                return Properties.Resources.BB_AddressesRng_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("bf10a3fc-7cae-46c5-b30e-757c3f5c90c7"); }
        }
    }
}