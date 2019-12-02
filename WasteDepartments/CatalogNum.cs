using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WasteDepartments
{
    class CatalogNum
    {
        public string CatalogNumber { get; set; }
        public string CatalogNumSon { get; set; }
        public string Description { get; set; }
        public int Department { get; set; }
        public string Machine { get; set; }
        public int Shift { get; set; }
        public int WorkCenter { get; set; }
        public string TableType { get; set; }

        public CatalogNum(string CatalogNumber, string CatalogNumSon, string Description,int Department,string Machine,int Shift,int WorkCenter,string TableType)
        {
            this.CatalogNumber = CatalogNumber;
            this.CatalogNumSon = CatalogNumSon;
            this.Description = Description;
            this.Department = Department;
            this.Machine = Machine;
            this.Shift = Shift;
            this.WorkCenter = WorkCenter;
            this.TableType = TableType;
        }

        public CatalogNum()
        {

        }
    }
}
