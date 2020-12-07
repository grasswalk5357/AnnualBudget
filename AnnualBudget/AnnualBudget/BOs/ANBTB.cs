using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTB : ERP_Common
    {
        private decimal tb001 = 0;
        private string tb002 = "";
        private string tb003 = "";
        private string tb004 = "";
        private string tb005 = "";
        private string tb006 = "";

        public ANBTB(string deptNo, string MainDeptNo, string year, string annualBudgetFormID, string annualBudgetFormName, string userId, string ModiDate) {
            this.Creator = userId;
            this.Usr_group = deptNo;
            this.Modifier = userId;
            this.Modi_date = ModiDate;
            this.Tb002 = MainDeptNo;
            this.Tb003 = year;
            this.Tb004 = annualBudgetFormID;
            this.Tb005 = annualBudgetFormName;
            
        }

        public decimal Tb001 { get => tb001; set => tb001 = value; }
        public string Tb002 { get => tb002; set => tb002 = value; }
        public string Tb003 { get => tb003; set => tb003 = value; }
        public string Tb004 { get => tb004; set => tb004 = value; }
        public string Tb005 { get => tb005; set => tb005 = value; }
        public string Tb006 { get => tb006; set => tb006 = value; }
    }
}
