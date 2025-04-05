using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components.Objects.Shapes
{
    public class GH_Ex_Shp_SetFont : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Shp_SetFont class.
        /// </summary>
        public GH_Ex_Shp_SetFont()
          : base("Shape Font", "Shp Font",
              "Sets the Shape Font properties",
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
            pManager.AddTextParameter("Family Name", "F", "Font Family name", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddColourParameter("Color", "C", "Font color", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddNumberParameter("Size", "S", "Font size", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddIntegerParameter("Justification", "J", "Text justifications", GH_ParamAccess.item);
            pManager[4].Optional = true;
            pManager.AddBooleanParameter("Is Bold", "B", "Font Bold status", GH_ParamAccess.item);
            pManager[5].Optional = true;
            pManager.AddBooleanParameter("Is Italic", "I", "Font Italic status", GH_ParamAccess.item);
            pManager[6].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[4];
            foreach (Justification value in Enum.GetValues(typeof(Justification)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Shape.Name, Constants.Shape.NickName, Constants.Shape.Output, GH_ParamAccess.item);
            pManager.AddTextParameter("Family Name", "F", "Font Family name", GH_ParamAccess.item);
            pManager.AddColourParameter("Color", "C", "Font color", GH_ParamAccess.item);
            pManager.AddNumberParameter("Size", "S", "Font size", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Justification", "J", "Text justifications", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Is Bold", "B", "Font Bold status", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Is Italic", "I", "Font Italic status", GH_ParamAccess.item);
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

            string family = "Arial";
            if (DA.GetData(1, ref family)) shape.FontFamily = family;

            Sd.Color color = Sd.Color.Black;
            if (DA.GetData(2, ref color)) shape.FontColor = color;

            double size = 10.0;
            if (DA.GetData(3, ref size)) shape.FontSize = size;

            int justifications = 1;
            if (DA.GetData(4, ref justifications)) shape.FontJustification = (Justification)justifications;

            bool isBold = false;
            if (DA.GetData(5, ref isBold)) shape.Bold = isBold;

            bool isItalic = false;
            if (DA.GetData(6, ref isItalic)) shape.Italic = isItalic;

            DA.SetData(0, shape);
            DA.SetData(1, shape.FontFamily);
            DA.SetData(2, shape.FontColor);
            DA.SetData(3, shape.FontSize);
            DA.SetData(4, shape.FontJustification);
            DA.SetData(5, shape.Bold);
            DA.SetData(6, shape.Italic);
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
                return Properties.Resources.BB_Shape_Font;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("b8d71103-f83d-490d-84fd-f6765a8c935d"); }
        }
    }
}