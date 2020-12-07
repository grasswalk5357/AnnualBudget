using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class CapEx : ERP_Common
    {
        private decimal no = 0;         // 序號
        private decimal id = 0;         // 唯一序號
        private string year = "";
        private string annualBudgetFormID = "";
        private string assetType = "";  // 資產類別
        private string name = "";       // 設備名稱
        private string spec = "";       // 規格
        private decimal num = 0;        // 數量
        private decimal unitPrice = 0;  // 單價
        private decimal totalPrice = 0; // 總價金額
        private decimal life = 0;       // 耐用年限
        private decimal getYear = 0;    // 預定取得(年)
        private decimal getMonth = 0;   // 預定取得(月)
        private decimal getDay = 0;     // 預定取得(日)
        private decimal depre = 0;      // 每月折舊
        private string purpose = "";    // 增設(改善)目的
        private string benifit = "";    // 預期效益評估
        private string trade = "";      // 交易對象
        private string isDelete = "";   // 是否標記為刪除

        public decimal No { get => no; set => no = value; }
        
        public string Name { get => name; set => name = value; }
        public string Spec { get => spec; set => spec = value; }
        public decimal Num { get => num; set => num = value; }
        public decimal UnitPrice { get => unitPrice; set => unitPrice = value; }
        public decimal TotalPrice { get => totalPrice; set => totalPrice = value; }
        public decimal Life { get => life; set => life = value; }
        public decimal GetYear { get => getYear; set => getYear = value; }
        public decimal GetMonth { get => getMonth; set => getMonth = value; }
        public decimal Depre { get => depre; set => depre = value; }
        public string Purpose { get => purpose; set => purpose = value; }
        public string Benifit { get => benifit; set => benifit = value; }
        public string Trade { get => trade; set => trade = value; }
        public decimal GetDay { get => getDay; set => getDay = value; }
        public string Year { get => year; set => year = value; }
        public string AnnualBudgetFormID { get => annualBudgetFormID; set => annualBudgetFormID = value; }
        public string IsDelete { get => isDelete; set => isDelete = value; }
        public string AssetType { get => assetType; set => assetType = value; }
        public decimal Id { get => id; set => id = value; }
    }
}
