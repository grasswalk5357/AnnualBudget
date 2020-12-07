using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTJ : ERP_Common
    {
        private decimal tj001 = 0;
        private string tj002 = "";
        private string tj003 = "";
        private decimal tj004 = 0;
        private string tj005 = "";

        public decimal Tj001 { get => tj001; set => tj001 = value; }
        public string Tj002 { get => tj002; set => tj002 = value; }
        public string Tj003 { get => tj003; set => tj003 = value; }
        public decimal Tj004 { get => tj004; set => tj004 = value; }
        public string Tj005 { get => tj005; set => tj005 = value; }
    }
}
