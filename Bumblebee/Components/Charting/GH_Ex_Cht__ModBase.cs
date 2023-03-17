using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public abstract class GH_Ex_Cht__ModBase : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Ch_ModBase class.
        /// </summary>
        public GH_Ex_Cht__ModBase()
          : base("GH_Ex_Ch_ModBase", "Nickname",
              "Description",
              "Category", "Subcategory")
        {
        }

        public GH_Ex_Cht__ModBase(string Name, string NickName, string Description, string Category, string Subcategory) : base(Name, NickName, Description, Category, Subcategory)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Chart.Name, Constants.Chart.NickName, Constants.Chart.Input, GH_ParamAccess.item);
            pManager.AddBooleanParameter("By Series", "B", "If true, values are plotted by series otherwise colors will be by point", GH_ParamAccess.item, false);
            pManager[1].Optional = true;

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Chart.Name, Constants.Chart.NickName, Constants.Chart.Output, GH_ParamAccess.item);
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
            get { return new Guid("9bf7456b-25b2-4da8-8e5b-4a93543d99f9"); }
        }
    }
}