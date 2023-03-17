using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using Sd = System.Drawing;

namespace Bumblebee.Components
{
    public class GH_Ex_Cht_SetGraphics : GH_Ex_Cht__ModBase
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Ch_SetGraphics class.
        /// </summary>
        public GH_Ex_Cht_SetGraphics()
          : base("Chart Graphics", "Chart Graphics",
              "Update Chart Graphics",
              Constants.ShortName, Constants.SubChart)
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
            base.RegisterInputParams(pManager);
            pManager.AddColourParameter("Fill Colors", "F", "Chart fill colors", GH_ParamAccess.list);
            pManager[2].Optional = true;
            pManager.AddColourParameter("Stroke Colors", "S", "Chart stroke colors", GH_ParamAccess.list);
            pManager[3].Optional = true;
            pManager.AddIntegerParameter("Stroke Weights", "W", "Chart stroke weights 0-3", GH_ParamAccess.list);
            pManager[4].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[4];
            foreach (BorderWeight value in Enum.GetValues(typeof(BorderWeight)))
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
            ExChart chart = null;
            if (!DA.GetData(0, ref chart)) return;
            chart = new ExChart(chart);

            bool mode = false;
            if (!DA.GetData(1, ref mode)) return;

            List<Sd.Color> fills = new List<Sd.Color>();
            if (DA.GetDataList(2, fills)) chart.SetFillColors(mode, fills);

            List<Sd.Color> strokes = new List<Sd.Color>();
            if (DA.GetDataList(3, strokes)) chart.SetStrokeColors(mode, strokes);

            List<int> weights = new List<int>();
            if (DA.GetDataList(4, weights)) chart.SetStrokeWeights(mode, weights);

            DA.SetData(0, chart);
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
                return Properties.Resources.BB_Chart_Graphics_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("1a2acfee-d57f-4c3c-b36d-307cc3a01cb6"); }
        }
    }
}