using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace Bumblebee.Components
{
    public class GH_Ex_Merge : GH_Ex_Rng_Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Merge class.
        /// </summary>
        public GH_Ex_Merge()
          : base("Range Merge", "Rng Mrg",
              "Merge the specified cells",
              Constants.ShortName, Constants.SubRange)
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
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager.AddBooleanParameter("Activate", "_A", "If true, the component will be activated", GH_ParamAccess.item, false);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddTextParameter("Start Address", "A", "The starting cell address of the range", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Status", "S", "Returns the status of the activate input", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;
            ExWorksheet worksheet = goo.ToWorksheet();

            IGH_Goo gooR = null;
            if (!DA.GetData(1, ref gooR)) return;
            ExRange range = new ExRange();
            if (!gooR.TryGetRange(ref range)) return;

            bool activate = false;
            DA.GetData(2, ref activate);

            if(activate) if(!range.IsSingle) worksheet.MergeCells(range);

            DA.SetData(0, worksheet);
            DA.SetData(1, range);
            DA.SetData(2, activate);
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
                return Properties.Resources.BB_Range_Merge3_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("d44031ec-dc80-48f8-b7ad-88c35856bdcf"); }
        }
    }
}