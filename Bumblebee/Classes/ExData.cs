using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = Microsoft.Office.Interop.Excel;

namespace Bumblebee
{
    public class ExData
    {
        #region members

        protected Dictionary<string,string> values = new Dictionary<string, string>();
        protected Dictionary<string, string> formats = new Dictionary<string, string>();

        #endregion

        #region constructors

        public ExData()
        {

        }

        public ExData(ExData exData)
        {
            foreach(string key in exData.values.Keys)
            {
                values.Add(key, values[key]);
                formats.Add(key, formats[key]);
            }
        }

        public ExData(List<string> columnNames, List<string> rowValues)
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

        public ExData(List<string> columnNames, List<string> rowValues, List<string> formatValues)
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
            return "Data | d:"+values.Count;
        }

        #endregion

    }
}
