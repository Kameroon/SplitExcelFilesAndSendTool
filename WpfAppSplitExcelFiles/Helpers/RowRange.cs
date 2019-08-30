using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SplitExcelFiles
{
    class RowRange
    {
        private int begin;
        private int end;

        public RowRange(int begin)
        {
            this.begin = begin;
            this.end = begin;
        }

        public int Begin
        {
            get { return begin; }
            set { begin = value; }
        }
        public int End
        {
            get { return end; }
            set { end = value; }
        }
    }
}
