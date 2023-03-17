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

        public static string SubRange
        {
            get { return "Range"; }
        }

        public static string SubObject
        {
            get { return "Shapes"; }
        }

        public static string SubCell
        {
            get { return "Cell"; }
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

        public static string SubChart
        {
            get { return "Charting"; }
        }

        public static string SubAnalysis
        {
            get { return "Analysis"; }
        }

        public static string SubWorkBooks
        {
            get { return "Workbooks"; }
        }

        public static string SubWorkSheets
        {
            get { return "WorkSheets"; }
        }

        public static Descriptor App
        {
            get { return new Descriptor("Excel Application", "App", "A Worksheet, Workbook, Range Object, Excel Application, or Text Workbook Name", "An Excel Application Object", "Excel Application Objects"); }
        }

        public static Descriptor Workbook
        {
            get { return new Descriptor("Workbook", "Wbk", "A Worksheet, Workbook, Range Object, Excel Application, or Text Workbook Name", "An Excel Workbook Object", "Excel Workbook Objects"); }
        }

        public static Descriptor Worksheet
        {
            get { return new Descriptor("Worksheet", "Wks", "A Worksheet, Workbook, Range Object, Excel Application, or Text Worksheet Name", "An Excel Worksheet Object", "Excel Worksheet Objects"); }
        }

        public static Descriptor Range
        {
            get { return new Descriptor("Range", "Rng", "A Range Object or Text Address (ex. A1:B1)", "An Excel Range Object", "Excel Range Objects"); }
        }

        public static Descriptor Cell
        {
            get { return new Descriptor("Cell", "Cel", "A Cell Object or Text Address (ex. A1)", "An Excel Cell Object", "Excel Cell Objects"); }
        }

        public static Descriptor Shape
        {
            get { return new Descriptor("Shape", "Shp", "A Smart Art, Control, or Illustrator Shape object", "An Excel Shape Object", "Excel Shape Objects"); }
        }

        public static Descriptor Chart
        {
            get { return new Descriptor("Chart", "Cht", "A Chart object", "A Chart object", "Chart objects"); }
        }

        public static Descriptor ShapeType
        {
            get { return new Descriptor("Shape Type", "T", "The Shape type", "The Shape type", "The Shape types"); }
        }

        public static Descriptor Boundary
        {
            get { return new Descriptor("Boundary", "B", "The Shape bounding rectangle", "The Shape bounding rectangle", "The Shape bounding rectangles"); }
        }

        public static Descriptor Activate
        {
            get { return new Descriptor("Activate", "_A", "If true, the component will be activated", "The active status of the component", "The active statuses of the component"); }
        }

        public static Descriptor DataSet
        {
            get { return new Descriptor("Data Set", "Dst", "A Data Set object", "A Data Set object", "Data Set objects"); }
        }

        #endregion

    }

    public class Descriptor
    {
        private string name = string.Empty;
        private string nickname = string.Empty;
        private string input = string.Empty;
        private string output = string.Empty;
        private string outputs = string.Empty;

        public Descriptor(string name, string nickname, string input, string output, string outputs)
        {
            this.name = name;
            this.nickname = nickname;
            this.input = input;
            this.output = output;
            this.outputs = outputs;
        }

        public virtual string Name
        {
            get { return name; }
        }

        public virtual string NickName
        {
            get { return nickname; }
        }

        public virtual string Input
        {
            get { return input; }
        }

        public virtual string Output
        {
            get { return output; }
        }

        public virtual string Outputs
        {
            get { return outputs; }
        }
    }
}
