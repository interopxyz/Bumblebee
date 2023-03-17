using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Cel_Location : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Location class.
        /// </summary>
        public GH_Ex_Cel_Location()
          : base("Cell Location", "XL Location",
              "Get the 1 based Row and Column index values for a cell",
              Constants.ShortName, Constants.SubCell)
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
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("Column", "C", "Column Index", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Row", "R", "Row Index", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooC = null;
            DA.GetData(0, ref gooC);
            ExCell cell = new ExCell();
            if (!gooC.TryGetCell(ref cell)) return;

            DA.SetData(0, cell.Column);
           DA.SetData(1, cell.Row);
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
                return Properties.Resources.BB_Location_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("1fd5c3f9-bbc1-421c-8698-6bbbeabc5505"); }
        }
    }
}