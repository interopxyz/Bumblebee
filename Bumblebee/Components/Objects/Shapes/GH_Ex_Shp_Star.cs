using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;


namespace Bumblebee.Components
{
    public class GH_Ex_Shp_Star : GH_Ex_Shp__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Sh_Star class.
        /// </summary>
        public GH_Ex_Shp_Star()
          : base("XL Star", "Star",
              "Adds a Star Shape object",
              Constants.ShortName, Constants.SubObject)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.quarternary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter(Constants.ShapeType.Name, Constants.ShapeType.NickName, Constants.ShapeType.Input, GH_ParamAccess.item, 0);
            pManager[3].Optional = true;
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item, false);
            pManager[4].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[3];
            foreach (ShapeStar value in Enum.GetValues(typeof(ShapeStar)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooS = null;
            if (!DA.GetData(0, ref gooS)) return;
            ExWorksheet worksheet = null;
            if (!gooS.TryGetWorksheet(ref worksheet)) return;

            string name = GetInstanceName();
            bool hasName = DA.GetData(1, ref name);

            Rectangle3d rect = GetBoundary(300, 100, 50, 50);
            DA.GetData(2, ref rect);

            int type = 0;
            DA.GetData(3, ref type);

            bool activate = false;
            DA.GetData(4, ref activate);

            if (activate) DA.SetData(1, worksheet.AddShape(name, (ShapeStar)type, rect));

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
                return Properties.Resources.BB_Shape_Star_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("d906009a-4313-470e-8df9-a47ad00e1861"); }
        }
    }
}