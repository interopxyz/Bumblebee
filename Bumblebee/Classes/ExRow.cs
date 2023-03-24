using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bumblebee
{
    public class ExRow:ExData
    {
        #region members

        protected Dictionary<string,string> values = new Dictionary<string, string>();
        protected Dictionary<string, string> formats = new Dictionary<string, string>();

        #endregion

        #region constructors

        public ExRow(ExRow exRow):base(exRow)
        {
            foreach(string key in exRow.values.Keys)
            {
                values.Add(key, values[key]);
                formats.Add(key, formats[key]);
            }
        }

        public ExRow(List<string> columnNames, List<string> rowValues) : base(DataTypes.Row)
        {
            int count = columnNames.Count;
            int vCount = rowValues.Count;
            for(int i = vCount;i<count;i++)
            {
                rowValues.Add("");
            }

            for(int i = 0; i < count; i++)
            {
                values.Add(columnNames[i],rowValues[i]);
                formats.Add(columnNames[i], "General");
            }
        }

        public ExRow(List<string> columnNames, List<string> rowValues, List<string> formatValues) : base(DataTypes.Row)
        {
            int count = columnNames.Count;
            int vCount = rowValues.Count;
            int fCount = formatValues.Count;

            for (int i = vCount; i < count; i++)
            {
                rowValues.Add("");
            }

            for (int i = fCount; i < count; i++)
            {
                formatValues.Add("General");
            }

            for (int i = 0; i < count; i++)
            {
                values.Add(columnNames[i], rowValues[i]);
                formats.Add(columnNames[i], formatValues[i]);
            }
        }

        #endregion

        #region properties

        public List<string> Columns
        {
            get { return values.Keys.ToList(); }
        }

        public List<string> Values
        {
            get { return values.Values.ToList(); }
        }

        public List<string> Formats
        {
            get { return formats.Values.ToList(); }
        }

        #endregion

        #region methods

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Row | c:"+values.Count;
        }

        #endregion

    }
}
