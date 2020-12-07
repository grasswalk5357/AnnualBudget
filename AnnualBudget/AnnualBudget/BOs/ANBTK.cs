using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTK : ERP_Common
    {
        private string tk001 = "";  // 流水號
        private string tk002 = "";  // 部門編號
        private string tk003 = "";  // 年度
        private string tk004 = "";  // 會科樣版編號
        private string tk005 = "";  // 會科樣版版本號
        private string tk006 = "";  // 刪除標記

        public string Tk001 { get => tk001; set => tk001 = value; }
        public string Tk002 { get => tk002; set => tk002 = value; }
        public string Tk003 { get => tk003; set => tk003 = value; }
        public string Tk004 { get => tk004; set => tk004 = value; }
        public string Tk005 { get => tk005; set => tk005 = value; }
        public string Tk006 { get => tk006; set => tk006 = value; }
    }
}
