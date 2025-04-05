using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace Bumblebee
{
    public class BumblebeeInfo : GH_AssemblyInfo
    {
        public override string Name
        {
            get
            {
                return "Bumblebee";
            }
        }
        public override Bitmap Icon
        {
            get
            {
                //Return a 24x24 pixel bitmap to represent this GHA library.
                return Properties.Resources.Bumblebee_Logo_24;
            }
        }
        public override string Description
        {
            get
            {
                //Return a short string describing the purpose of this GHA library.
                return "Excel interoperability plugin for Grasshopper 3d";
            }
        }
        public override Guid Id
        {
            get
            {
                return new Guid("23df6520-1dee-431b-b2d4-492d3ee91c7a");
            }
        }

        public override string AuthorName
        {
            get
            {
                //Return a string identifying you or your company.
                return "David Mans";
            }
        }
        public override string AuthorContact
        {
            get
            {
                //Return a string representing your preferred contact details.
                return "interopxyz@gmail.com";
            }
        }

        public override string AssemblyVersion
        {
            get
            {
                return "3.0.5.0";
            }
        }
    }

    public class BitmapPlusCategoryIcon : GH_AssemblyPriority
    {
        public object Properties { get; private set; }

        public override GH_LoadingInstruction PriorityLoad()
        {
            Instances.ComponentServer.AddCategoryIcon(Constants.ShortName, Bumblebee.Properties.Resources.BB_TabLogo_C_01);
            Instances.ComponentServer.AddCategorySymbolName(Constants.ShortName, 'B');
            return GH_LoadingInstruction.Proceed;
        }
    }
}
