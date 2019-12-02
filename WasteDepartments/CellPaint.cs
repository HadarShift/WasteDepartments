using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WasteDepartments
{
    /// <summary>
    /// שומר את מיקומי התאים שצריכים להיצבע
    /// </summary>
    class CellPaint
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public CellPaint(int row,int column)
        {
            Row = row;
            Column = column;
        }
    }
}
