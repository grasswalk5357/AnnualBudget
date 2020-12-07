using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class BizTrip : ERP_Common
    {
        private decimal id = 0;                 // 唯一序號
        private string year = "";
        private string annualBudgetFormID = "";
        private decimal no = 0;                 // 序號
        private decimal month = 0;              // 月份
        private string name = "";               // 出差者
        private string location = "";           // 出差地區
        private decimal days = 0;               // 天數
        private decimal airFare = 0;            // 機票款
        private decimal hotelFare = 0;          // 住宿費
        private decimal shippingExpenses = 0;   // 交通費
        private decimal otherFare = 0;          // 雜費
        private decimal dailyExpense = 0;       // 日支費
        private decimal foodStipend = 0;        // 餐費
        private decimal sumOfTripExpense = 0;   // 旅費小計
        private decimal entertainment = 0;      // 交際費
        private decimal tripInsurance = 0;      // 旅平險
        private string memo = "";               // 出差目的說明
        private string isDelete = "";   // 是否標記為刪除


        public decimal Month { get => month; set => month = value; }
        public string Name { get => name; set => name = value; }
        public string Location { get => location; set => location = value; }
        public decimal Days { get => days; set => days = value; }
        public decimal AirFare { get => airFare; set => airFare = value; }
        public decimal HotelFare { get => hotelFare; set => hotelFare = value; }
        public decimal ShippingExpenses { get => shippingExpenses; set => shippingExpenses = value; }
        public decimal OtherFare { get => otherFare; set => otherFare = value; }
        public decimal DailyExpense { get => dailyExpense; set => dailyExpense = value; }
        public decimal FoodStipend { get => foodStipend; set => foodStipend = value; }
        public decimal SumOfTripExpense { get => sumOfTripExpense; set => sumOfTripExpense = value; }
        public decimal Entertainment { get => entertainment; set => entertainment = value; }
        public decimal TripInsurance { get => tripInsurance; set => tripInsurance = value; }
        public decimal No { get => no; set => no = value; }
        public string Memo { get => memo; set => memo = value; }
        public string Year { get => year; set => year = value; }
        public string AnnualBudgetFormID { get => annualBudgetFormID; set => annualBudgetFormID = value; }
        public string IsDelete { get => isDelete; set => isDelete = value; }
        public decimal Id { get => id; set => id = value; }
    }
}
