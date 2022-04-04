using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components.Application
{
    public class GH_Ex_Location : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Location class.
        /// </summary>
        public GH_Ex_Location()
          : base("Cell Location", "XL Location",
              "Description",
              Constants.ShortName, Constants.SubApp)
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
            pManager.AddTextParameter("Address", "A", "The resulting cell address", GH_ParamAccess.item,"A1");
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
            string address = "A1";
            if(!DA.GetData(0, ref address))return;

            Tuple<int, int> location = Helper.GetCellLocation(address);

           DA.SetData(0, location.Item1-1);
           DA.SetData(1, location.Item2-1);
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