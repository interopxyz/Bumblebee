using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Dst_WriteFast : GH_Ex_Wks__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_FastWrite class.
        /// </summary>
        public GH_Ex_Dst_WriteFast()
          : base("Fast Write Data", "XL Fast",
              "Fast Write data to Excel",
              Constants.ShortName, Constants.SubData)
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
            pManager[0].Optional = true;
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Input, GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddTextParameter("Values", "V", "A datatree of values", GH_ParamAccess.tree);
            pManager.AddBooleanParameter("Clear", "_X", "If true, all contents of the sheet will be cleared prior to writing new data", GH_ParamAccess.item, false);
            pManager[3].Optional = true;
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddGenericParameter(Constants.Range.Name, Constants.Range.NickName, Constants.Range.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool clear = false;
            DA.GetData(3, ref clear);

            bool active = false;
            if (!DA.GetData(4, ref active)) return;

            if (active)
            {
                IGH_Goo goo = null;
                DA.GetData(0, ref goo);
                ExWorksheet worksheet = null;
                if (!goo.TryGetWorksheet(ref worksheet)) return;

                IGH_Goo gooC = null;
                DA.GetData(1, ref gooC);
                ExCell cell = new ExCell();
                gooC.TryGetCell(ref cell);

                GH_Structure<GH_String> ghData = new GH_Structure<GH_String> ();

                List<List<GH_String>> dataSet = new List<List<GH_String>>();
                if(!DA.GetDataTree(2, out ghData))return;

                foreach (List<GH_String> data in ghData.Branches)
                {
                    dataSet.Add(data);
                }

                if (clear) worksheet.ClearSheet();
                ExRange range = worksheet.WriteData(dataSet, cell);

                DA.SetData(0, worksheet);
                DA.SetData(1, range);
            }
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
                return Properties.Resources.BB_Sheet_Write_Fast_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("9e022976-510f-429f-a9d8-f659415168ef"); }
        }
    }
}