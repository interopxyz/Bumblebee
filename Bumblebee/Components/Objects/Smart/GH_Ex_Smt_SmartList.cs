using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Smt_SmartList : GH_Ex_Smt__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Sm_SmartList class.
        /// </summary>
        public GH_Ex_Smt_SmartList()
          : base("Smart Art List", "Smart List",
              "Create a Smart Art list object",
              Constants.ShortName, Constants.SubObject)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.hidden; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);

            Param_Integer paramA = (Param_Integer)pManager[5];
            foreach (BarChartType value in Enum.GetValues(typeof(BarChartType)))
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
            DA.GetData(0, ref gooS);
            ExWorksheet worksheet = null;
            bool hasWs = gooS.TryGetWorksheet(ref worksheet);

            string name = GetInstanceName();
            DA.GetData(1, ref name);

            Rectangle3d boundary = new Rectangle3d(Plane.WorldXY, new Point3d(250, 250, 0), new Point3d(500, 500, 0));
            DA.GetData(2, ref boundary);

            List<string> values = new List<string>();
            DA.GetDataList(3, values);

            List<int> levels = new List<int>();
            DA.GetDataList(4, levels);

            int type = 0;
            DA.GetData(5, ref type);

            bool activate = false;
            DA.GetData(6, ref activate);

            if (activate)
            {
                worksheet.Freeze();
                ExShape smart = new ExShape(worksheet, name, boundary,values,levels);
                smart.SetList(type);
                worksheet.UnFreeze();

                DA.SetData(1, smart);
            }

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
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("674571ad-fecb-4e07-8d4b-8359d42d3dc2"); }
        }
    }
}