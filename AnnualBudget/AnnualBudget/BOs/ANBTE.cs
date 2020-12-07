using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTE
    {
        private decimal te000 = 0;  // 流水號
        private string te001 = "";  // 單別代號
        private string te002 = "";  // 部門代號
        private string te003 = "";  // 年度

        private decimal te004 = 0;  // 月份
        private string te005 = "";  // 出差者
        private string te006 = "";  // 出差地區
        private decimal te007 = 0;  // 天數
        private decimal te008 = 0;  // 機票款
        private decimal te009 = 0;  // 住宿費
        private decimal te010 = 0;  // 交通費
        private decimal te011 = 0;  // 雜費
        private decimal te012 = 0;  // 日支費
        private decimal te013 = 0;  // 旅費
        private decimal te014 = 0;  // 旅費小計
        private decimal te015 = 0;  // 交際費
        private decimal te016 = 0;  // 旅平險
        private string te017 = "";  // 出差目的說明
        private string te018 = "";  // 是否標記刪除

        public decimal Te000 { get => te000; set => te000 = value; }
        public string Te001 { get => te001; set => te001 = value; }
        public string Te002 { get => te002; set => te002 = value; }
        public string Te003 { get => te003; set => te003 = value; }
        public decimal Te004 { get => te004; set => te004 = value; }
        public string Te005 { get => te005; set => te005 = value; }
        public string Te006 { get => te006; set => te006 = value; }
        public decimal Te007 { get => te007; set => te007 = value; }
        public decimal Te008 { get => te008; set => te008 = value; }
        public decimal Te009 { get => te009; set => te009 = value; }
        public decimal Te010 { get => te010; set => te010 = value; }
        public decimal Te011 { get => te011; set => te011 = value; }
        public decimal Te012 { get => te012; set => te012 = value; }
        public decimal Te013 { get => te013; set => te013 = value; }
        public decimal Te014 { get => te014; set => te014 = value; }
        public decimal Te015 { get => te015; set => te015 = value; }
        public decimal Te016 { get => te016; set => te016 = value; }
        public string Te017 { get => te017; set => te017 = value; }
        public string Te018 { get => te018; set => te018 = value; }
    }
}
