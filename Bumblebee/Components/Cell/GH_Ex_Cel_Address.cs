using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Cel_Address : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Address class.
        /// </summary>
        public GH_Ex_Cel_Address()
          : base("Cell Address", "XL Address",
              "Creates a Cell from a column and row index",
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
            pManager.AddIntegerParameter("Column", "C", "Column Index", GH_ParamAccess.item, 1);
            pManager[0].Optional = true;
            pManager.AddIntegerParameter("Row", "R", "Row Index", GH_ParamAccess.item, 1);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Absolute Column", "AC", "Set absolute value for column", GH_ParamAccess.item, false);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Absolute Row", "AR", "Set absolute value for row", GH_ParamAccess.item, false);
            pManager[3].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            int col = 1;
            DA.GetData(0, ref col);

            int row = 1;
            DA.GetData(1, ref row);

            bool absC = false;
            DA.GetData(2, ref absC);

            bool absR = false;
            DA.GetData(3, ref absR);

            DA.SetData(0, new ExCell(col,row,absC,absR));
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
                return Properties.Resources.BB_Address_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("a922d90f-a323-478a-af0c-2782431179e3"); }
        }
    }
}