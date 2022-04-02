using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.IO;

namespace Bumblebee.Components
{
    public class GH_Ex_Wb_Save : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Wb_Save class.
        /// </summary>
        public GH_Ex_Wb_Save()
          : base("Save Workbook", "Save Workbook",
              "Save a workbook to a .xlxs file",
              Constants.ShortName, "Workbook")
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
            pManager.AddGenericParameter("Workbook", "Wb", "The Excel Workbook object", GH_ParamAccess.item);
            pManager.AddTextParameter("Folder Path", "F", "The path to the workbook", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddTextParameter("File Name", "N", "The workbook name", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Save", "S", "If true, the workbook will be saved", GH_ParamAccess.item);
            pManager[3].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Workbook", "Wb", "The Excel Workbook object", GH_ParamAccess.item);
            pManager.AddTextParameter("FilePath", "P", "The full filepath to the saved file", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {

            ExWorkbook workbook = null;
            if (!DA.GetData<ExWorkbook>(0, ref workbook)) return;

            string path = "C:\\Users\\Public\\Documents\\";
            bool hasPath = DA.GetData(1, ref path);

            if (!hasPath)
            {
                if (this.OnPingDocument().FilePath != null)
                {
                    path = Path.GetDirectoryName(this.OnPingDocument().FilePath) + "\\";
                }
            }

            if (!Directory.Exists(path))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "The file provided path does not exist. Please verify this is a valid file path.");
                return;
            }

            string name = DateTime.UtcNow.ToString("yyyy-dd-M_HH-mm-ss");
            DA.GetData(2, ref name);

            bool save = false;
            DA.GetData(3, ref save);

            string filepath = path + name + ".xlsx";
            if (save)
            {
                workbook.Save(filepath);
                DA.SetData(1, filepath);
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
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("491b3f78-97ac-47ad-9b52-2cbf7e38318c"); }
        }
    }
}