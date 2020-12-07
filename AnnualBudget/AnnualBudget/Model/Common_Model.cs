using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AnnualBudget.Model
{
    class Common_Model
    {
		public static List<Object> LoadMatrixData(string annualBudgetFormID, string deptNo, string year, string TmplID)
		{

			List<Object> list = null;

			if (annualBudgetFormID.Equals("ABM001"))
				list = ANBTC_Model.LoadData(annualBudgetFormID, deptNo, year);
			
			else if (annualBudgetFormID.Equals("ABM002"))
				list = ANBTD_Model.LoadData(annualBudgetFormID, deptNo, year);
			
			else if (annualBudgetFormID.Equals("ABM003"))
				list = ANBTE_Model.LoadData(annualBudgetFormID, deptNo, year);
			
			else if (annualBudgetFormID.Equals("ABM004"))
				list = ANBTF_Model.LoadData(annualBudgetFormID, deptNo, year);
			
			else if (annualBudgetFormID.Equals("ABM007"))  
				list = ANBTI_Model.LoadData(annualBudgetFormID, deptNo, year, TmplID);

			else if (annualBudgetFormID.Equals("ABM008"))
				list = ANBTN_Model.LoadData(annualBudgetFormID, deptNo, year);


			return list;
			
		}


		public static DataGridView SetDGV(string annualBudgetFormID, DataGridView dgv, List<Object> list, DataTable tmplTable, bool isDeptAccounting, string Dept)
		{
			if (annualBudgetFormID.Equals("ABM001"))
				dgv = ANBTC_Model.SetDGV(dgv, list);

			else if (annualBudgetFormID.Equals("ABM002"))
				dgv = ANBTD_Model.SetDGV(dgv, list);
			
			else if (annualBudgetFormID.Equals("ABM003"))
				dgv = ANBTE_Model.SetDGV(dgv, list);
			
			else if (annualBudgetFormID.Equals("ABM004"))
				dgv = ANBTF_Model.SetDGV(dgv, list);
			
			else if (annualBudgetFormID.Equals("ABM007"))
				dgv = ANBTI_Model.SetDGV(dgv, list, tmplTable, isDeptAccounting, Dept);
			
			else if (annualBudgetFormID.Equals("ABM008"))
				dgv = ANBTN_Model.SetDGV(dgv, list);


			return dgv;
		}
	}
}
