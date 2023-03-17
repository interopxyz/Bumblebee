using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Cel_Addresses : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Addresses class.
        /// </summary>
        public GH_Ex_Cel_Addresses()
          : base("Generate Addresses", "Addresses",
              "Generate a list of Cells from a start and extent Cell.",
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
            pManager.AddGenericParameter("Start Cell", "A", "The starting cell of the range", GH_ParamAccess.item);
            pManager.AddGenericParameter("Extent Cell", "B", "The cell at the extent of the range", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Flip", "F", "If true, cells are listed by column. If false, by row", GH_ParamAccess.item, true);
            pManager[2].Optional = true;

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Outputs, GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {

            IGH_Goo gooCa = null;
            DA.GetData(0, ref gooCa);
            ExCell a = new ExCell();
            if (!gooCa.TryGetCell(ref a)) return;

            IGH_Goo gooCb = null;
            DA.GetData(1, ref gooCb);
            ExCell b = new ExCell();
            if (!gooCb.TryGetCell(ref b)) return;

            bool flip = true;
            DA.GetData(2, ref flip);
            
            int xa = Math.Min(a.Column, b.Column);
            int xb = Math.Max(a.Column, b.Column)+1;

            int ya = Math.Min(a.Row, b.Row);
            int yb = Math.Max(a.Row, b.Row)+1;

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

            DA.SetDataList(0, cells);
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
                return Properties.Resources.BB_Addresses2_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("b2449c36-a03a-41e4-93ef-1849ff5ab9ae"); }
        }
    }
}