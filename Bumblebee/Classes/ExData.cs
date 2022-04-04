using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bumblebee
{
    public abstract class ExData
    {

        #region members

        public enum DataTypes { Row,Column}
        protected DataTypes dataType = DataTypes.Row;

        #endregion

        #region constructors

        public ExData(ExData exData)
        {
            this.dataType = exData.dataType;
        }

        protected ExData(DataTypes dataType)
        {
            this.dataType = dataType;
        }

        #endregion

        #region properties

        public virtual DataTypes DataType
        {
            get { return dataType; }
        }

        #endregion

        #region methods



        #endregion

        #region members



        #endregion

    }
}
