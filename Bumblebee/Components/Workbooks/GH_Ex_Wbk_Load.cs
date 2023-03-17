using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.IO;

namespace Bumblebee.Components
{
    public class GH_Ex_Wbk_Load : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Wb_Load class.
        /// </summary>
        public GH_Ex_Wbk_Load()
          : base("Load Workbook", "Load Workbook",
              "Load a Workbook from a filepath.",
              Constants.ShortName, Constants.SubWorkBooks)
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
            pManager.AddGenericParameter(Constants.App.Name, Constants.App.NickName, Constants.App.Input, GH_ParamAccess.item);
            pManager.AddTextParameter("File Path", "P", "The name of an active Workbook", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Workbook.Name, Constants.Workbook.NickName, Constants.Workbook.Output, GH_ParamAccess.item);
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

            string name = string.Empty;
            if (!DA.GetData(1, ref name)) return;

            if(!File.Exists(name))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "The file provided path does not exist. Please verify this is a valid file path.");
                return;
            }

            ExWorkbook workbook = app.LoadWorkbook(name);

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
                return Properties.Resources.BB_Book_Load_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("ca5fead8-0a4a-46bc-9b8c-178477198957"); }
        }
    }
}