using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class HumanBudget : ERP_Common
    {
        public HumanBudget() {
            //MonthlyData monthlyData = new MonthlyData();
            this.monthlyData = new decimal[13];
        }
        private string year = "";
        private string annualBudgetFormID = "";
        private decimal id = 0;         // 唯一識別號
        private decimal no = 0;
        //private string deptNo = "";
        
        private string role = "";       // 直接間接人員
        private string jobTitle = "";   // 職稱
        private string rank = "";   // 職等
        private string jobContent = ""; // 工作職掌
        private decimal actNum = 0;     // 實際人數
        private decimal estNum = 0;     // 計劃人數
        private decimal startMonth = 0; // 增減起始月份
        private decimal diffNum = 0;  // 增減人數
        private decimal salary = 0;     // 增減之人員每月薪資
        private decimal totalChangeSalary = 0;  // 每月增減薪資
        private string reason = "";     // 增減人力原因
        private decimal[] monthlyData;
        private string isDelete = "";   // 是否標記為刪除
        private string refCln = "";     // 參照欄位

        public decimal No { get => no; set => no = value; }
        public string Role { get => role; set => role = value; }
        public string JobTitle { get => jobTitle; set => jobTitle = value; }
        public string Rank { get => rank; set => rank = value; }
        public string JobContent { get => jobContent; set => jobContent = value; }
        public decimal ActNum { get => actNum; set => actNum = value; }
        public decimal EstNum { get => estNum; set => estNum = value; }
        public decimal StartMonth { get => startMonth; set => startMonth = value; }
        public decimal DiffNum { get => diffNum; set => diffNum = value; }
        public decimal Salary { get => salary; set => salary = value; }
        public decimal TotalChangeSalary { get => totalChangeSalary; set => totalChangeSalary = value; }
        public string Reason { get => reason; set => reason = value; }
        public decimal[] MonthlyData { get => monthlyData; set => monthlyData = value; }
        public string AnnualBudgetFormID { get => annualBudgetFormID; set => annualBudgetFormID = value; }
        //public string DeptNo { get => deptNo; set => deptNo = value; }
        public string Year { get => year; set => year = value; }
        public string IsDelete { get => isDelete; set => isDelete = value; }
        public decimal Id { get => id; set => id = value; }
        
        public string RefCln { get => refCln; set => refCln = value; }
    }
}
