﻿using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Wbk_GetAll : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Wb_GetAll class.
        /// </summary>
        public GH_Ex_Wbk_GetAll()
          : base("Get Workbooks", "Get Workbooks",
              "Gets all the currently active Workbooks",
              Constants.ShortName, Constants.SubWorkBooks)
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
            pManager.AddGenericParameter(Constants.App.Name, Constants.App.NickName, Constants.App.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Workbook.Name, Constants.Workbook.NickName, Constants.Workbook.Outputs, GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            DA.GetData(0, ref gooA);
            ExApp app = null;
            if (!gooA.TryGetApp(ref app)) return;

            List<ExWorkbook> workbooks = app.GetWorkbooks();

            DA.SetDataList(0, workbooks);
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
                return Properties.Resources.BB_Books_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("8a6d8f6b-9a5b-49eb-a738-845f7724e8dd"); }
        }
    }
}