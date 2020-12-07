using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTD
    {
        private decimal td000 = 0;  // 流水號
        private string td001 = "";  // 單別代號
        private string td002 = "";  // 部門代號
        private string td003 = "";  // 年度
        private string td004 = "";  // 職務能力
        private string td005 = "";  // 順練重點
        private string td006 = "";  // 對象
        private decimal td007 = 0;  // 人數
        private string td008 = "";  // 內外訓
        private decimal td009 = 0;  // 起始月
        private decimal td010 = 0;  // 結束月
        private decimal td011 = 0;  // 時數
        private decimal td012 = 0;  // 每月費用
        private decimal td013 = 0;  // 該計劃費用總額
        private string td014 = "";  // 備註
        private string td015 = "";  // 是否標記刪除

        public decimal Td000 { get => td000; set => td000 = value; }
        public string Td001 { get => td001; set => td001 = value; }
        public string Td002 { get => td002; set => td002 = value; }
        public string Td003 { get => td003; set => td003 = value; }
        public string Td004 { get => td004; set => td004 = value; }
        public string Td005 { get => td005; set => td005 = value; }
        public string Td006 { get => td006; set => td006 = value; }
        public decimal Td007 { get => td007; set => td007 = value; }
        public string Td008 { get => td008; set => td008 = value; }
        public decimal Td009 { get => td009; set => td009 = value; }
        public decimal Td010 { get => td010; set => td010 = value; }
        public decimal Td011 { get => td011; set => td011 = value; }
        public decimal Td012 { get => td012; set => td012 = value; }
        public decimal Td013 { get => td013; set => td013 = value; }
        public string Td014 { get => td014; set => td014 = value; }
        public string Td015 { get => td015; set => td015 = value; }
    }
}
