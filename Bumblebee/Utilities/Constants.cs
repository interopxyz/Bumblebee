using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bumblebee
{
    public class Constants
    {

        #region naming

        public static string LongName
        {
            get { return ShortName + " v" + Major + "." + Minor; }
        }

        public static string ShortName
        {
            get { return "Bumblebee2"; }
        }

        private static string Minor
        {
            get { return typeof(Constants).Assembly.GetName().Version.Minor.ToString(); }
        }
        private static string Major
        {
            get { return typeof(Constants).Assembly.GetName().Version.Major.ToString(); }
        }

        public static string SubApp
        {
            get { return "App"; }
        }

        public static string SubData
        {
            get { return "Data"; }
        }

        public static string SubGraphics
        {
            get { return "Graphics"; }
        }

        public static string SubWorkBooks
        {
            get { return "Workbooks"; }
        }

        public static string SubWorkSheets
        {
            get { return "WorkSheets"; }
        }

        #endregion

    }
}
