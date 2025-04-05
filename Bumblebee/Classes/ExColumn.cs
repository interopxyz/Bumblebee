using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bumblebee
{
    public class ExColumn:ExData
    {

        #region members

        protected string name = string.Empty;
        protected string format = "General";

        protected List<string> values = new List<string>();

        #endregion

        #region constructors

        public ExColumn(ExColumn exColumn) : base(exColumn)
        {
            this.name = exColumn.name;
            this.format = exColumn.format;
            this.values = exColumn.values;
        }

        public ExColumn(List<string> values) : base(DataTypes.Column)
        {
            this.values = values;
        }

        public ExColumn(string name, List<string> values) : base(DataTypes.Column)
        {
            this.name = name;
            this.values = values;
        }

        public ExColumn(string name, List<string> values, string format) : base(DataTypes.Column)
        {
            this.name = name;
            this.format = format;
            this.values = values;
        }

        public ExColumn(List<string> values, string format) : base(DataTypes.Column)
        {
            this.format = format;
            this.values = values;
        }

        #endregion

        #region properties

        public virtual string Name
        {
            get { return name; }
        }

        public virtual string Format
        {
            get { return format; }
        }

        public virtual bool IsNumeric
        {
            get
            {
                foreach(string value in values)
                {
                    if (!double.TryParse(value, out double num)) return false;
                }
                return true;
            }
        }

        public virtual bool IsFormula
        {
            get
            {
                foreach (string value in values)
                {
                    if (!value.ToCharArray()[0].Equals('=')) return false;
                }
                return true;
            }
        }

        public virtual List<string> Values
        {
            get { return values; }
        }

        #endregion

        #region methods



        #endregion

        #region overrides

        public override string ToString()
        {
            return "Column | r:" + values.Count + " ("+format+")";
        }

        #endregion

    }
}
