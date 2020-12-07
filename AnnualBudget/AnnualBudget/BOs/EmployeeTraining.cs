using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class EmployeeTraining : ERP_Common
    {
        private decimal no = 0;
        private decimal id = 0;         // 唯一識別號
        private string annualBudgetFormID = "";
        private string year = "";
        private string ability = "";
        private string focalPoint = "";
        private string target = "";
        private decimal peopleNum = 0;
        private string in_or_Ex = "";
        private decimal startMonth = 0;
        private decimal endMonth = 0;
        private decimal hours = 0;
        private decimal monthlyFee = 0;
        private decimal projectQuota = 0;
        private string memo = "";
        private decimal[] monthlyData;
        private string isDelete = "";   // 是否標記為刪除

        public decimal No { get => no; set => no = value; }
        public string Ability { get => ability; set => ability = value; }
        public string FocalPoint { get => focalPoint; set => focalPoint = value; }
        public string Target { get => target; set => target = value; }
        public decimal PeopleNum { get => peopleNum; set => peopleNum = value; }
        public string In_or_Ex { get => in_or_Ex; set => in_or_Ex = value; }
        public decimal StartMonth { get => startMonth; set => startMonth = value; }
        public decimal EndMonth { get => endMonth; set => endMonth = value; }
        public decimal Hours { get => hours; set => hours = value; }
        public decimal MonthlyFee { get => monthlyFee; set => monthlyFee = value; }
        public decimal ProjectQuota { get => projectQuota; set => projectQuota = value; }
        public string Memo { get => memo; set => memo = value; }
        public string IsDelete { get => isDelete; set => isDelete = value; }
        public string Year { get => year; set => year = value; }
        public string AnnualBudgetFormID { get => annualBudgetFormID; set => annualBudgetFormID = value; }
        public decimal Id { get => id; set => id = value; }
    }
}
