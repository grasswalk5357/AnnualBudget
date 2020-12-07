using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTO : ERP_Common
    {
        private decimal to001 = 0;  // 唯一序號
        private string to002 = "";  // 工號
        private string to003 = "";  // 姓名
        private string to004 = "";  // AD帳號
        private string to005 = "";  // 權限是否禁止

        public decimal To001 { get => to001; set => to001 = value; }
        public string To002 { get => to002; set => to002 = value; }
        public string To003 { get => to003; set => to003 = value; }
        public string To004 { get => to004; set => to004 = value; }
        public string To005 { get => to005; set => to005 = value; }
    }
}
