using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTH : ERP_Common
    {
        public ANBTH() { }

        public ANBTH(string rdp_id)
        {
            this.Th004 = rdp_id;
        }

        private decimal th000 = 0;  // 流水號
        private string th001 = "";  // 單別代號
        private string th002 = "";  // 部門代號
        private string th003 = "";  // 年度


        private string th004 = "";   // 專案序號
        private string th005 = "";   // 細項序號 1-9
        private string th006 = "";   // 細項名稱
        private decimal[] monthlyData = new decimal[13];    // 月份資料

        private decimal th019 = 0;   // 總金額
        private decimal th020 = 0;   // 明躍支出
        private decimal th021 = 0;   // 立和支出
        private decimal th022 = 0;   // 客戶支出
        private string th023 = "";   // 是否標記為刪除
        private decimal th024 = 0;   // 每月折舊

        public decimal Th000 { get => th000; set => th000 = value; }
        public string Th001 { get => th001; set => th001 = value; }
        public string Th002 { get => th002; set => th002 = value; }
        public string Th003 { get => th003; set => th003 = value; }
        public string Th004 { get => th004; set => th004 = value; }
        public string Th005 { get => th005; set => th005 = value; }
        public string Th006 { get => th006; set => th006 = value; }
        public decimal[] MonthlyData { get => monthlyData; set => monthlyData = value; }
        public decimal Th019 { get => th019; set => th019 = value; }
        public decimal Th020 { get => th020; set => th020 = value; }
        public decimal Th021 { get => th021; set => th021 = value; }
        public decimal Th022 { get => th022; set => th022 = value; }
        public string Th023 { get => th023; set => th023 = value; }
        public decimal Th024 { get => th024; set => th024 = value; }
    }
}
