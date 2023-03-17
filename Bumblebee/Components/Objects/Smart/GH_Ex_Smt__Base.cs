using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public abstract class GH_Ex_Smt__Base : GH_Ex_Wks__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Sm_Base class.
        /// </summary>
        public GH_Ex_Smt__Base()
          : base("GH_Ex_Ch_Base", "Nickname",
              "Description",
              "Category", "Subcategory")
        {
        }

        public GH_Ex_Smt__Base(string Name, string NickName, string Description, string Category, string Subcategory) : base(Name, NickName, Description, Category, Subcategory)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager.AddTextParameter("Name", "N", "The title of the chart", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddRectangleParameter("Boundary", "B", "The chart frame", GH_ParamAccess.item, new Rectangle3d(Plane.WorldXY, new Point3d(250, 250, 0), new Point3d(500, 500, 0)));
            pManager[2].Optional = true;
            pManager.AddTextParameter("Values", "V", "Input list for data to be displayed in smart art formatting", GH_ParamAccess.list);
            pManager.AddIntegerParameter("Levels", "L", "Corresponding list of indices which specify the level at which the data will be displayed", GH_ParamAccess.list);
            pManager.AddIntegerParameter("Type", "T", "The Smart Object type", GH_ParamAccess.item);
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item, false);
            pManager[6].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddGenericParameter(Constants.Shape.Name, Constants.Shape.NickName, Constants.Shape.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
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
            get { return new Guid("18f61f76-382f-4d4c-b257-90b3b01f7e72"); }
        }
    }
}