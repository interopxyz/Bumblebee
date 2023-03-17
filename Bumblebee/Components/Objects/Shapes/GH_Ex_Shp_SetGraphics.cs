using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components
{
    public class GH_Ex_Shp_SetGraphics : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Ch_SetGraphics class.
        /// </summary>
        public GH_Ex_Shp_SetGraphics()
          : base("Shape Graphics", "Shape Graphics",
              "Update Shape Graphics",
              Constants.ShortName, Constants.SubObject)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.quinary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Shape.Name, Constants.Shape.NickName, Constants.Shape.Input, GH_ParamAccess.item);
            pManager.AddColourParameter("Fill Color", "F", "Shape fill color", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddColourParameter("Stroke Color", "S", "Shape stroke color", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddNumberParameter("Stroke Weight", "W", "Shape stroke weight", GH_ParamAccess.item);
            pManager[3].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Shape.Name, Constants.Shape.NickName, Constants.Shape.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            ExShape shape = null;
            if (!DA.GetData(0, ref shape)) return;
            shape = new ExShape(shape);

            Sd.Color fill = Sd.Color.White;
            if (DA.GetData(1,ref fill)) shape.SetFillColor(fill);

            Sd.Color stroke = Sd.Color.Black;
            if (DA.GetData(2, ref stroke)) shape.SetStrokeColor(stroke);

            double weight = 1.0;
            if (DA.GetData(3,ref weight)) shape.SetStrokeWeight(weight);

            DA.SetData(0, shape);
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
                return Properties.Resources.BB_Shape_Graphics_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("c9dda40f-c9c8-4a0b-ad0f-be2e9f7c4e45"); }
        }
    }
}