using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTI : ERP_Common
    {
        private decimal ti000 = 0;  // 流水號
        private string ti001 = "";  // 單別代號
        private string ti002 = "";  // 部門代號
        private string ti003 = "";  // 年度

        private string ti004 = "";  // 會科樣版號
        private string ti005 = "";  // 會科樣版版本號
        private decimal ti006 = 0;  // 樣版科目流水號
        private string ti007 = "";  // 科目編號
        private string ti008 = "";  // 科目名稱
        private string ti009 = "";  // 科目說明
        private decimal ti010 = 0;  // 月份1
        private decimal ti011 = 0;  // 月份2
        private decimal ti012 = 0;  // 月份3
        private decimal ti013 = 0;  // 月份4
        private decimal ti014 = 0;  // 月份5
        private decimal ti015 = 0;  // 月份6
        private decimal ti016 = 0;  // 月份7
        private decimal ti017 = 0;  // 月份8
        private decimal ti018 = 0;  // 月份9
        private decimal ti019 = 0;  // 月份10
        private decimal ti020 = 0;  // 月份11
        private decimal ti021 = 0;  // 月份12
        private decimal ti022 = 0;  // 項目總計
        private decimal ti023 = 0;  // 預留
        private string ti024 = "";  // 預留

        public decimal Ti000 { get => ti000; set => ti000 = value; }
        public string Ti001 { get => ti001; set => ti001 = value; }
        public string Ti002 { get => ti002; set => ti002 = value; }
        public string Ti003 { get => ti003; set => ti003 = value; }
        public string Ti004 { get => ti004; set => ti004 = value; }
        public string Ti005 { get => ti005; set => ti005 = value; }
        public decimal Ti006 { get => ti006; set => ti006 = value; }
        public string Ti007 { get => ti007; set => ti007 = value; }
        public string Ti008 { get => ti008; set => ti008 = value; }
        public string Ti009 { get => ti009; set => ti009 = value; }
        public decimal Ti010 { get => ti010; set => ti010 = value; }
        public decimal Ti011 { get => ti011; set => ti011 = value; }
        public decimal Ti012 { get => ti012; set => ti012 = value; }
        public decimal Ti013 { get => ti013; set => ti013 = value; }
        public decimal Ti014 { get => ti014; set => ti014 = value; }
        public decimal Ti015 { get => ti015; set => ti015 = value; }
        public decimal Ti016 { get => ti016; set => ti016 = value; }
        public decimal Ti017 { get => ti017; set => ti017 = value; }
        public decimal Ti018 { get => ti018; set => ti018 = value; }
        public decimal Ti019 { get => ti019; set => ti019 = value; }
        public decimal Ti020 { get => ti020; set => ti020 = value; }
        public decimal Ti021 { get => ti021; set => ti021 = value; }
        public decimal Ti022 { get => ti022; set => ti022 = value; }
        public decimal Ti023 { get => ti023; set => ti023 = value; }
        public string Ti024 { get => ti024; set => ti024 = value; }
    }
}
