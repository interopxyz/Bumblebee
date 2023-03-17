using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public abstract class GH_Ex_Wks__Base : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Ws_Base class.
        /// </summary>
        public GH_Ex_Wks__Base()
          : base("GH_Ex_Ws_Base", "Nickname",
              "Description",
              "Category", "Subcategory")
        {
        }

        public GH_Ex_Wks__Base(string Name, string NickName, string Description, string Category, string Subcategory) : base(Name, NickName, Description, Category, Subcategory)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Worksheet.Name, Constants.Worksheet.NickName, Constants.Worksheet.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Worksheet.Name, Constants.Worksheet.NickName, Constants.Worksheet.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
        }

        protected Rectangle3d GetBoundary(int x, int y, int w, int h)
        {
            return new Rectangle3d(Plane.WorldXY, new Point3d(x, y, 0), new Point3d(x + w, y + h, 0));
        }

        protected string GetInstanceName()
        {
            return this.InstanceGuid.ToString() + "-" + this.RunCount;
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
            get { return new Guid("bc7817b1-7cc1-473e-847b-1753f9ad2e4e"); }
        }
    }
}