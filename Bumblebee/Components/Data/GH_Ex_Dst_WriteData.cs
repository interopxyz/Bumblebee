using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Dt_WriteData : GH_Ex_Wks__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_WriteData class.
        /// </summary>
        public GH_Ex_Dt_WriteData()
          : base("Write Data", "XL Write",
              "Write data to Excel",
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
            pManager.AddGenericParameter(Constants.DataSet.Name, Constants.DataSet.NickName, Constants.DataSet.Input, GH_ParamAccess.list);
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

                List<ExData> genData = new List<ExData>();
                if (!DA.GetDataList(2, genData)) return;

                List<ExRow> rows = new List<ExRow>();
                List<ExColumn> cols = new List<ExColumn>();

                foreach (ExData data in genData)
                {
                    if (data.DataType == ExData.DataTypes.Column)
                    {
                        cols.Add((ExColumn)data);
                    }
                    else
                    {
                        rows.Add((ExRow)data);
                    }
                }

                if ((cols.Count > 0) & (rows.Count > 0)) return;
                if (clear) worksheet.ClearSheet();
                ExRange range = new ExRange();
                if (cols.Count > 0) range = worksheet.WriteData(cols, cell);
                if (rows.Count > 0) range = worksheet.WriteData(rows, cell);

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
                return Properties.Resources.BB_Sheet_Write_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("b372e027-76eb-4ec8-a49c-76fcb2f7985b"); }
        }
    }
}