using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

namespace Bumblebee
{
    public class ExRange
    {
        protected ExCell start = new ExCell();
        protected ExCell extent = new ExCell();
        #region members


        #endregion

        #region constructors

        public ExRange()
        {

        }

        public ExRange(ExCell start, ExCell extent)
        {
            this.Start = start;
            this.Extent = extent;
        }

        public ExRange(string range)
        {
            string[] cells = range.Split(':');
            this.start = new ExCell(cells[0]);
            if (cells.Count() > 1)
            {
                this.extent = new ExCell(cells[1]);
            }
            else
            {
                this.extent = new ExCell(cells[0]);
            }

        }

        public ExRange(ExRange range)
        {
            this.start = range.Start;
            this.extent = range.Extent;
        }

        #endregion

        #region properties

        public virtual ExCell Start
        {
            get { return new ExCell(start); }
            set { start = new ExCell(value); }
        }

        public virtual ExCell Extent
        {
            get { return new ExCell(extent); }
            set { extent = new ExCell(value); }
        }

        public virtual bool IsSingle
        {
            get { return ((start.Column == extent.Column) & (start.Row == extent.Row)); }
        }

        #endregion

        #region methods



        #endregion

        #region overrides

        public override string ToString()
        {
            return this.start.Address+":"+this.extent.Address;
        }

        #endregion

    }
}
