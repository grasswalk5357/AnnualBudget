using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTN : ERP_Common
    {
        private decimal tn000 = 0;  // 流水號
        private string tn001 = "";  // 單別代號
        private string tn002 = "";  // 部門代號
        private string tn003 = "";  // 年度

        private string tn004 = "";  // 部門(單位)
        private string tn005 = "";  // 職稱/職位
        private string tn006 = "";  // 職等職級
        private decimal tn007 = 0;  // 實際人數
        private string tn008 = "";  // 現職者
        private decimal tn009 = 0;  // 計畫人數
        private decimal tn010 = 0;  // 計畫人數與實際人數差額
        private decimal tn011 = 0;  // 增減起始月份

        private string tn012 = "";  // 增減人力原因
        private string tn013 = "";  // 是否標記刪除
        private string tn014 = "";  // 備用欄位

        public decimal Tn000 { get => tn000; set => tn000 = value; }
        public string Tn001 { get => tn001; set => tn001 = value; }
        public string Tn002 { get => tn002; set => tn002 = value; }
        public string Tn003 { get => tn003; set => tn003 = value; }
        public string Tn004 { get => tn004; set => tn004 = value; }
        public string Tn005 { get => tn005; set => tn005 = value; }
        public string Tn006 { get => tn006; set => tn006 = value; }
        public decimal Tn007 { get => tn007; set => tn007 = value; }
        public string Tn008 { get => tn008; set => tn008 = value; }
        public decimal Tn009 { get => tn009; set => tn009 = value; }
        public decimal Tn010 { get => tn010; set => tn010 = value; }
        public decimal Tn011 { get => tn011; set => tn011 = value; }        
        public string Tn012 { get => tn012; set => tn012 = value; }
        public string Tn013 { get => tn013; set => tn013 = value; }
        public string Tn014 { get => tn014; set => tn014 = value; }
    }
}
