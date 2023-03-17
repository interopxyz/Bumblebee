﻿using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Ctl_ListBox : GH_Ex_Ctl__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Ct_ListBox class.
        /// </summary>
        public GH_Ex_Ctl_ListBox()
          : base("XL List Box", "ListBox",
              "Adds a List Box Control Shape object",
              Constants.ShortName, Constants.SubObject)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.tertiary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Input, GH_ParamAccess.item);
            pManager.AddTextParameter("Options", "O", "The dropdown options", GH_ParamAccess.list);
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item, false);
            pManager[5].Optional = true;
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
            DA.GetData(1, ref name);

            Rectangle3d boundary = GetBoundary(10, 250, 100, 50);
            DA.GetData(2, ref boundary);

            IGH_Goo gooC = null;
            if (!DA.GetData(3, ref gooC)) return;
            ExCell cell = null;
            if (!gooC.TryGetCell(ref cell)) return;

            List<string> data = new List<string>();
            if (!DA.GetDataList(4, data)) return;

            bool activate = false;
            DA.GetData(5, ref activate);

            if (activate) DA.SetData(1, worksheet.AddListBox(name, cell, data, boundary));

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
                return Properties.Resources.BB_Controls_ListBox_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("017e0093-43de-4c54-a33f-de18ae8ec568"); }
        }
    }
}