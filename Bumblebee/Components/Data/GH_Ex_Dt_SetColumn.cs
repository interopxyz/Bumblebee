using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components.Data
{
    public class GH_Ex_Dt_SetColumn : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Dt_Column class.
        /// </summary>
        public GH_Ex_Dt_SetColumn()
          : base("Compile Column", "Col",
              "Compile data into column assemblies",
              Constants.ShortName, Constants.SubData)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Column Name", "N", "The names of the column", GH_ParamAccess.item);
            pManager.AddTextParameter("Row Values", "R", "The row values corresponding to each column", GH_ParamAccess.list);
            pManager.AddTextParameter("Format", "F", "A MS Office Number Format" +
                Environment.NewLine + "Examples (\"General\", \"hh: mm:ss\", \"$#,##0.0\" " +
                Environment.NewLine + "https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=en-us&rs=en-us&ad=us", GH_ParamAccess.item);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("DataSet", "Ds", "A compiled DataSet", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string name = string.Empty;
            if (!DA.GetData(0, ref name)) return;

            List<string> values = new List<string>();
            if (!DA.GetDataList(1, values)) return;

            string formats = string.Empty;
            bool hasFormat = DA.GetData(2,ref formats);

            ExColumn data;

            if (hasFormat)
            {
                data = new ExColumn(name, values, formats);
            }
            else
            {
                data = new ExColumn(name, values);
            }

            DA.SetData(0, data);
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
                return Properties.Resources.BB_Columns_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("914a0693-2477-4753-9937-248affead56f"); }
        }
    }
}