using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTF
    {
        private decimal tf000 = 0;  // 流水號
        private string tf001 = "";  // 單別代號
        private string tf002 = "";  // 部門代號
        private string tf003 = "";  // 年度

        private string tf004 = "";  // 資產類別
        private string tf005 = "";  // 設備名稱
        private string tf006 = "";  // 規格
        private decimal tf007 = 0;  // 數量
        private decimal tf008 = 0;  // 單價
        private decimal tf009 = 0;  // 總價金額
        private decimal tf010 = 0;  // 耐用年限
        private decimal tf011 = 0;  // 預定取得-年
        private decimal tf012 = 0;  // 預定取得-月
        private decimal tf013 = 0;  // 預定取得-日
        private decimal tf014 = 0;  // 每月折舊
        private string tf015 = "";  // 增設改善目的
        private string tf016 = "";  // 預期效益評估
        private string tf017 = "";  // 交易對像
        private string tf018 = "";  // 是否標記刪除

        public decimal Tf000 { get => tf000; set => tf000 = value; }
        public string Tf001 { get => tf001; set => tf001 = value; }
        public string Tf002 { get => tf002; set => tf002 = value; }
        public string Tf003 { get => tf003; set => tf003 = value; }
        public string Tf004 { get => tf004; set => tf004 = value; }
        public string Tf005 { get => tf005; set => tf005 = value; }
        public string Tf006 { get => tf006; set => tf006 = value; }
        public decimal Tf007 { get => tf007; set => tf007 = value; }
        public decimal Tf008 { get => tf008; set => tf008 = value; }
        public decimal Tf009 { get => tf009; set => tf009 = value; }
        public decimal Tf010 { get => tf010; set => tf010 = value; }
        public decimal Tf011 { get => tf011; set => tf011 = value; }
        public decimal Tf012 { get => tf012; set => tf012 = value; }
        public decimal Tf013 { get => tf013; set => tf013 = value; }
        public decimal Tf014 { get => tf014; set => tf014 = value; }
        public string Tf015 { get => tf015; set => tf015 = value; }
        public string Tf016 { get => tf016; set => tf016 = value; }
        public string Tf017 { get => tf017; set => tf017 = value; }
        public string Tf018 { get => tf018; set => tf018 = value; }
    }
}
