using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Wbk_RunMacro : GH_Ex_Wbk__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Wb_RunMacro class.
        /// </summary>
        public GH_Ex_Wbk_RunMacro()
          : base("Run Macro", "Run Macro",
              "Run a macro in a Workbook",
              Constants.ShortName, Constants.SubWorkBooks)
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
            pManager.AddTextParameter("Name", "N", "The unique name of the Macro", GH_ParamAccess.item);
            pManager.AddGenericParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item);
            pManager[2].Optional = true;
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
            IGH_Goo gooB = null;
            if (!DA.GetData(0, ref gooB)) return;
            ExWorkbook workbook = null;
            if (!gooB.TryGetWorkbook(ref workbook)) return;

            string name = string.Empty;
            if (!DA.GetData(1, ref name)) return;

            bool activate = false;
            DA.GetData(2, ref activate);


            if (activate)
            {
                workbook.RunMacro(name);
            }

            DA.SetData(0, workbook);
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
                return Properties.Resources.BB_Book_RunMacro2_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("eafa7f95-1be8-4c47-b120-5520807e360e"); }
        }
    }
}