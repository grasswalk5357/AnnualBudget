using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    public class ERP_Common
    {
        public ERP_Common() { }

		private string company = "WFTW05";
		private string creator = "ITADMIN";
		private string usr_group = "110";
		private string create_date = DateTime.Now.ToString("yyyyMMdd");
		private string modifier = "";
		private string modi_date = "";
		private double flag = 1;
		private string deptNo = "";

		public string Company { get => company; set => company = value; }
		public string Creator { get => creator; set => creator = value; }
		public string Usr_group { get => usr_group; set => usr_group = value; }
		public string Create_date { get => create_date; set => create_date = value; }
		public string Modifier { get => modifier; set => modifier = value; }
		public string Modi_date { get => modi_date; set => modi_date = value; }
		public double Flag { get => flag; set => flag = value; }
        public string DeptNo { get => deptNo; set => deptNo = value; }
    }
}
