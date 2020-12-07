using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    public class UserInfo
    {
        private string userID = "";
        private string userName = "";
        private string deptNo = "";
        private string deptName = "";
        private string parentDeptNo = "";
        private string parentDeptName = "";
        private UserInfo parentDept = null;
        private string level = "";



        public string UserID { get => userID; set => userID = value; }
        public string DeptNo { get => deptNo; set => deptNo = value; }
        public string DeptName { get => deptName; set => deptName = value; }
        public string UserName { get => userName; set => userName = value; }
        public string ParentDeptNo { get => parentDeptNo; set => parentDeptNo = value; }
        public string ParentDeptName { get => parentDeptName; set => parentDeptName = value; }
        public UserInfo ParentDept { get => parentDept; set => parentDept = value; }
        public string Level { get => level; set => level = value; }
    }
}
