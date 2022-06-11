using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components.Objects
{
    public class GH_Ex_ObjPicture : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_ObjPicture class.
        /// </summary>
        public GH_Ex_ObjPicture()
          : base("Picture", "Pic",
              "A Picture object",
              Constants.ShortName, Constants.SubObject)
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
            pManager.AddGenericParameter("Worksheet / Workbook / App", "Ws", "A Workbook, Worksheet, or Excel Application", GH_ParamAccess.item);
            pManager.AddTextParameter("Filepath", "P", "A filepath to an image", GH_ParamAccess.item);
            pManager.AddPointParameter("Location", "L", "A pixel based location for the image", GH_ParamAccess.item, new Point3d(100, 100, 0));
            pManager[2].Optional = true;
            pManager.AddNumberParameter("Scale", "S", "A scale value", GH_ParamAccess.item, 1.0);
            pManager[3].Optional = true;
            pManager.AddBooleanParameter("Activate", "_A", "If true, the component will be activated", GH_ParamAccess.item, false);
            pManager[4].Optional = true;
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
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;
            ExWorksheet worksheet = goo.ToWorksheet();

            string filepath = string.Empty;
            if (!DA.GetData(1, ref filepath)) return;

            Point3d location = new Point3d(100, 100, 0);
            DA.GetData(2, ref location);

            double scale = 1.0;
            DA.GetData(3, ref scale);

            bool activate = false;
            DA.GetData(4, ref activate);

            if(activate) worksheet.AddPicture(filepath, location.X, location.Y, scale);

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
                return Properties.Resources.BB_Obj_Picture1_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("c745425b-4dc5-45cc-9238-2940f1b126e2"); }
        }
    }
}