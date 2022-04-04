using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bumblebee
{
    public class ExCell
    {

        #region members

        protected int column = 1;
        protected int row = 1;
        protected string address = "A1";
        protected bool isColumnAbsolute = false;
        protected bool isRowAbsolute = false;

        #endregion

        #region constructors

        public ExCell()
        {

        }

        public ExCell(ExCell cell)
        {
            this.column = cell.column;
            this.row = cell.row;
            this.address = cell.address;

            this.isColumnAbsolute = cell.isColumnAbsolute;
            this.isRowAbsolute = cell.isRowAbsolute;
        }

        public ExCell(string address)
        {
            this.address = address;
            Tuple<int, int> loc = Helper.GetCellLocation(address);
            Tuple<bool, bool> abs = Helper.GetCellAbsolute(address);

            this.column = loc.Item1;
            this.row = loc.Item2;

            this.isColumnAbsolute = abs.Item1;
            this.isRowAbsolute = abs.Item2;
        }

        public ExCell(int column, int row)
        {
            this.column = column;
            this.row = row;
            this.address = Helper.GetCellAddress(column, row);
        }

        public ExCell(int column, int row, bool absColumn, bool absRow)
        {
            this.column = column;
            this.row = row;
            this.address = Helper.GetCellAddress(column, row);
            this.isColumnAbsolute = absColumn;
            this.isRowAbsolute = absRow;
        }

        #endregion

        #region properties

        public virtual int Column
        {
            get { return column; }
        }

        public virtual int Row
        {
            get { return row; }
        }

        public virtual string Address
        {
            get { return address; }
        }

        public virtual bool IsColumnAbsolute
        {
            get { return isColumnAbsolute; }
        }

        public virtual bool IsRowAbsolute
        {
            get { return isRowAbsolute; }
        }


        #endregion

        #region methods



        #endregion

        #region overrides

        public override string ToString()
        {
            return "Range | " + this.Address;
        }

        #endregion

    }
}
