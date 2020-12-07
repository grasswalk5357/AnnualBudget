using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTC : ERP_Common
    {
        private decimal tc000 = 0;  // 流水號
        private string tc001 = "";  // 單別代號
        private string tc002 = "";  // 部門代號
        private string tc003 = "";  // 年度
        private string tc004 = "";  // 直接間接人員
        private string tc005 = "";  // 職稱
        private string tc006 = "";  // 職等
        private string tc007 = "";  // 工作職掌
        private decimal tc008 = 0;  // 實際人數
        private decimal tc009 = 0;  // 計劃人數
        private decimal tc010 = 0;  // 增減起始月份
        private decimal tc011 = 0;  // 增減人數
        private decimal tc012 = 0;  // 增減之人員每月薪資
        private decimal tc013 = 0;  // 每月增減薪資
        private string tc014 = "";  // 增減人力原因
        private string tc015 = "";  // 是否標記刪除
        private string tc016 = "";  // 參照號碼

        public decimal Tc000 { get => tc000; set => tc000 = value; }
        public string Tc001 { get => tc001; set => tc001 = value; }
        public string Tc002 { get => tc002; set => tc002 = value; }
        public string Tc003 { get => tc003; set => tc003 = value; }
        public string Tc004 { get => tc004; set => tc004 = value; }
        public string Tc005 { get => tc005; set => tc005 = value; }
        public string Tc006 { get => tc006; set => tc006 = value; }
        public string Tc007 { get => tc007; set => tc007 = value; }
        public decimal Tc008 { get => tc008; set => tc008 = value; }
        public decimal Tc009 { get => tc009; set => tc009 = value; }
        public decimal Tc010 { get => tc010; set => tc010 = value; }
        public decimal Tc011 { get => tc011; set => tc011 = value; }
        public decimal Tc012 { get => tc012; set => tc012 = value; }
        public decimal Tc013 { get => tc013; set => tc013 = value; }
        public string Tc014 { get => tc014; set => tc014 = value; }
        public string Tc015 { get => tc015; set => tc015 = value; }
        public string Tc016 { get => tc016; set => tc016 = value; }
    }
}
