using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components.Cell
{
    public class GH_Ex_Addresses : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Addresses class.
        /// </summary>
        public GH_Ex_Addresses()
          : base("Generate Addresses", "Addresses",
              "Generate a list of cell addresses from a start and extent cell.",
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
            pManager.AddTextParameter("Start Address", "A", "The starting cell address of the range", GH_ParamAccess.item);
            pManager.AddTextParameter("Extent Address", "B", "The cell address at the extent of the range", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Flip", "F", "If true, cells are listed by column. If false, by row", GH_ParamAccess.item, true);
            pManager[2].Optional = true;

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Addresses", "A", "The addresses in the range", GH_ParamAccess.list);
            pManager.AddIntegerParameter("Column Indices", "C", "The column indices", GH_ParamAccess.list);
            pManager.AddIntegerParameter("Row Indices", "R", "The row indices", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string a = "A1";
            if (!DA.GetData(0, ref a)) return;

            string b = "A1";
            if (!DA.GetData(1, ref b)) return;

            bool flip = true;
            DA.GetData(2, ref flip);

            Tuple<int, int> A = Helper.GetCellLocation(a);
            Tuple<int, int> B = Helper.GetCellLocation(b);

            int xa = Math.Min(A.Item1, B.Item1);
            int xb = Math.Max(A.Item1, B.Item1);

            int ya = Math.Min(A.Item2, B.Item2);
            int yb = Math.Max(A.Item2, B.Item2);

            List<string> addresses = new List<string>();
            List<int> columns = new List<int>();
            List<int> rows = new List<int>();

            if (flip)
            {
                for (int i = xa-1; i < xb; i++)
                {
                    for (int j = ya-1; j < yb; j++)
                    {
                        addresses.Add(Helper.GetCellAddress(i, j));
                        columns.Add(i);
                        rows.Add(j);
                    }
                }
            }
            else
            {
                for (int j = xa-1; j < xb; j++)
                {
                    for (int i = ya-1; i < yb; i++)
                    {
                        addresses.Add(Helper.GetCellAddress(i, j));
                        columns.Add(i);
                        rows.Add(j);
                    }
                }
            }

            DA.SetDataList(0, addresses);
            DA.SetDataList(1, columns);
            DA.SetDataList(2, rows);
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
            get { return new Guid("b2449c36-a03a-41e4-93ef-1849ff5ab9ae"); }
        }
    }
}