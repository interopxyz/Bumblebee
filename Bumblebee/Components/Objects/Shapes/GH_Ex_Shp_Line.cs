using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Shp_Line : GH_Ex_Shp__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Sh_Line class.
        /// </summary>
        public GH_Ex_Shp_Line()
          : base("XL Line", "Line",
              "Adds a Line Shape object",
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
            pManager.AddLineParameter("Line", "L", "A line", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Start Arrow", "S", "The start arrow type", GH_ParamAccess.item, 0);
            pManager[3].Optional = true;
            pManager.AddIntegerParameter("End Arrow", "E", "The end arrow type", GH_ParamAccess.item, 0);
            pManager[4].Optional = true;
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item, false);
            pManager[5].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[3];
            foreach (ArrowStyle value in Enum.GetValues(typeof(ArrowStyle)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramB = (Param_Integer)pManager[4];
            foreach (ArrowStyle value in Enum.GetValues(typeof(ArrowStyle)))
            {
                paramB.AddNamedValue(value.ToString(), (int)value);
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

            Line line = new Line();
            DA.GetData(2, ref line);

            int arrowStart = 0;
            DA.GetData(3, ref arrowStart);

            int arrowEnd = 0;
            DA.GetData(4, ref arrowEnd);

            bool activate = false;
            DA.GetData(5, ref activate);

            if (activate) DA.SetData(1, worksheet.AddLine(name,(ArrowStyle)arrowStart,(ArrowStyle)arrowEnd, line));

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
                return Properties.Resources.BB_Shape_Line_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("0f43cc44-d8ea-46bb-907f-78d81a19e8d6"); }
        }
    }
}