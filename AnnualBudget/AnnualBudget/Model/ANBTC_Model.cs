using AnnualBudget.BOs;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AnnualBudget.Model
{
    class ANBTC_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// 新增或更新資料至ANBTC資料表
		/// </summary>
		/// <param name="list">存放ANBTC物件的List</param>
		/// <returns></returns>
		public static bool UpdateToANBTC(List<Object> list)
		{
			
			bool result = true;
			StringBuilder SQL_1 = new StringBuilder();
			StringBuilder SQL_2 = new StringBuilder();
			EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
			SqlConnectionStringBuilder sqlsb;

			sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["AB_ConnString"].ConnectionString);

			sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

			using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
			{
				if (conn.State != ConnectionState.Open)
					conn.Open();

				SqlCommand command = conn.CreateCommand();
				SqlTransaction trans = conn.BeginTransaction(IsolationLevel.ReadCommitted);
				command.Connection = conn;
				command.Transaction = trans;

				try
				{
					foreach (HumanBudget myObj in list) { 
						SQL_1.Clear();
						SQL_2.Clear();

						SQL_1.AppendFormat("IF EXISTS (SELECT TC000 FROM dbo.ANBTC WHERE TC000 = {0}) " +
												"UPDATE dbo.ANBTC SET FLAG = FLAG + 1, MODIFIER = '{1}', MODI_DATE = '{2}', " +
												"TC005 = '{3}', TC006 = '{4}', TC007 = '{5}', TC008 = {6}, TC009 = {7}, " +
												"TC010 = {8}, TC011 = {9}, TC012 = {10}, TC013 = {11}, TC014 = '{12}', " +
												"TC015 = '{13}', TC016 = '{14}' " +
												"WHERE TC000 = {15} ELSE ",
												myObj.Id, myObj.Modifier, myObj.Modi_date, myObj.JobTitle, myObj.Rank, 
												myObj.JobContent, myObj.ActNum, myObj.EstNum, myObj.StartMonth, myObj.DiffNum, 
												myObj.Salary, myObj.TotalChangeSalary, myObj.Reason, myObj.IsDelete, myObj.RefCln, myObj.Id);


						SQL_2.AppendFormat("INSERT INTO dbo.ANBTC " +
							"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, " +
							"TC001, TC002, TC003, TC004, TC005, TC006, " +
							"TC007, TC008, TC009, TC010, TC011, TC012, " +
							"TC013, TC014, TC015, TC016 ) " +
							"VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, " +
							"'{7}', '{8}', '{9}', '{10}', '{11}', " +
							"'{12}', '{13}', {14}, {15}, {16}, " +
							"{17}, {18}, {19}, '{20}', '{21}', " +
							"'{22}') ",
							myObj.Company, myObj.Creator, myObj.Usr_group, myObj.Create_date, "", "", myObj.Flag,
							myObj.AnnualBudgetFormID, myObj.DeptNo, myObj.Year, myObj.Role, myObj.JobTitle, myObj.Rank, 
							myObj.JobContent, myObj.ActNum, myObj.EstNum, myObj.StartMonth, myObj.DiffNum, myObj.Salary,
							myObj.TotalChangeSalary, myObj.Reason, myObj.IsDelete, myObj.RefCln);

						SQL_1 = SQL_1.Append(SQL_2);

						
						command.CommandText = SQL_1.ToString();
						command.ExecuteNonQuery();
					}
					trans.Commit();

				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTC」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「ANBTC」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTC」時發生錯誤的訊息：" + e2.Message);
					}

					result = false;
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return result;
			}
		}


		/// <summary>
		/// 將人力需求預算表DataGridView的資料儲存至List
		/// </summary>
		/// <param name="dgv">人力需求預算表的DataGridView</param>
		/// <param name="AnnualBudgetFormID">人力需求預算表的代碼</param>
		/// <param name="mainDeptNo">編寫年度預算的目標部門代號</param>
		/// <param name="deptNo">登入者的部門代號</param>
		/// <param name="year">年份</param>
		/// <param name="userId">登入者的ID</param>
		/// <param name="modidate">修改日期</param>
		/// <returns></returns>
		public static List<Object> SaveMatrixData(DataGridView dgv, string AnnualBudgetFormID,string mainDeptNo, string deptNo, string year, string userId, string modidate)
        {
			List<Object> list = new List<Object>();
            HumanBudget myObject = null;

			try
			{
				for (int i = 0; i < dgv.RowCount - 1; i++)
				{
					myObject = new HumanBudget();

					myObject.Creator = userId;
					myObject.Usr_group = deptNo;

					myObject.Modifier = userId;
					myObject.Modi_date = modidate;

					myObject.DeptNo = mainDeptNo;
					myObject.Year = year;

					myObject.AnnualBudgetFormID = AnnualBudgetFormID;

					myObject.No = Convert.ToDecimal(dgv.Rows[i].HeaderCell.Value);

					if (dgv.Rows[i].Cells["dgv_HBT_cln_Role"].Value != null)
						myObject.Role = dgv.Rows[i].Cells["dgv_HBT_cln_Role"].Value.ToString(); // 直接/間接人員

					if (dgv.Rows[i].Cells["dgv_HBT_cln_JobTitle"].Value != null)
						myObject.JobTitle = dgv.Rows[i].Cells["dgv_HBT_cln_JobTitle"].Value.ToString();	// 職稱

					if (dgv.Rows[i].Cells["dgv_HBT_cln_Rank"].Value != null)
						myObject.Rank = dgv.Rows[i].Cells["dgv_HBT_cln_Rank"].Value.ToString();		// 職等

					if (dgv.Rows[i].Cells["dgv_HBT_cln_JobContent"].Value != null)
						myObject.JobContent = dgv.Rows[i].Cells["dgv_HBT_cln_JobContent"].Value.ToString();		// 工作職掌

					myObject.ActNum = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_HBT_cln_ActNum"].Value);			// 實際人數
					myObject.EstNum = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_HBT_cln_EstNum"].Value);			// 計劃人數
					myObject.StartMonth = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_HBT_cln_StartMonth"].Value); // 增減起始月份
					myObject.DiffNum = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_HBT_cln_DiffNum"].Value);       // 增減人數
					myObject.Salary = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_HBT_cln_Salary"].Value);         // 增減之人員每月薪資
					myObject.TotalChangeSalary = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_HBT_cln_TotalChangeSalary"].Value);   // 每月增減新資
					if (dgv.Rows[i].Cells["dgv_HBT_cln_Reason"].Value != null)
						myObject.Reason = dgv.Rows[i].Cells["dgv_HBT_cln_Reason"].Value.ToString();     // 增減人力原因

					DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dgv.Rows[i].Cells["dgv_HBT_cbc_Delete"];	// 刪除標記
					if (chk.Value != null && (bool)chk.Value)
						myObject.IsDelete = "Y";
					else
						myObject.IsDelete = "N";

					if (dgv.Rows[i].Cells["dgv_HBT_cln_Ref_ANBTN"].Value != null)
						myObject.RefCln = dgv.Rows[i].Cells["dgv_HBT_cln_Ref_ANBTN"].Value.ToString();     // 參照欄位
					

					myObject.Id = Convert.ToDecimal(dgv.Rows[i].Cells[12].Value);	// 唯一序號
					
					list.Add(myObject);
				}
			}
			
			catch (Exception e) {
				_log.Debug("儲存[人力需求預算表]時發生錯誤：" + e.Message);
				MessageBox.Show("儲存[人力需求預算表]時發生錯誤：" + e.Message);
			}


			return list;
        }


		/// <summary>
		/// 檢查DataGridView中的必填(帶**)欄位是否都已填
		/// </summary>
		/// <param name="dgv"></param>
		/// <returns></returns>
		public static bool CheckMatrixData(DataGridView dgv)
		{			
			HumanBudget obj = null;
			bool result = true;
			bool isDeleted = false;

			try
			{
				for (int i = 0; i < dgv.RowCount - 1; i++)
				{
					if (dgv.Rows[i].Visible == true) // 只針對Visible == true的row做處理
					{
						DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dgv.Rows[i].Cells["dgv_HBT_cbc_Delete"];   // 刪除標記
						if (chk.Value != null && (bool)chk.Value)
							isDeleted = true;
						else
							isDeleted = false;

						if (isDeleted == false)	// 只針對沒有被勾取的做處理 
						{ 
							obj = new HumanBudget();

							if (dgv.Rows[i].Cells["dgv_HBT_cln_Role"].Value != null)
								obj.Role = dgv.Rows[i].Cells["dgv_HBT_cln_Role"].Value.ToString(); // 直接/間接人員										

							if (dgv.Rows[i].Cells["dgv_HBT_cln_JobContent"].Value != null)
								obj.JobContent = dgv.Rows[i].Cells["dgv_HBT_cln_JobContent"].Value.ToString();     // 工作職掌

							obj.StartMonth = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_HBT_cln_StartMonth"].Value); // 增減起始月份

							obj.Salary = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_HBT_cln_Salary"].Value);         // 增減之人員每月薪資

							if (String.IsNullOrEmpty(obj.Role) || String.IsNullOrEmpty(obj.JobContent) ||
								String.IsNullOrEmpty(obj.StartMonth.ToString()) || String.IsNullOrEmpty(obj.Salary.ToString())) {						
								result = false;
								return result;
							}
						}
					}
				}
			}
			catch (Exception e)
			{
				_log.Debug("儲存[人力需求預算表]時發生錯誤：" + e.Message);
				MessageBox.Show("儲存[人力需求預算表]時發生錯誤：" + e.Message);
				result = false;
			}

			return result;
		}


		/// <summary>
		/// 依據必填欄位是否已填，變更該欄位的底色
		/// </summary>
		/// <param name="dgv"></param>
		public static void CheckDataValue(DataGridView dgv) {
			for (int i = 0; i < dgv.RowCount - 1; i++) { 
				if (dgv.Rows[i].Cells["dgv_HBT_cln_Role"].Value != null && !String.IsNullOrEmpty(dgv.Rows[i].Cells["dgv_HBT_cln_Role"].Value.ToString()))
					dgv.Rows[i].Cells["dgv_HBT_cln_Role"].Style.BackColor = Color.White;
				else
					dgv.Rows[i].Cells["dgv_HBT_cln_Role"].Style.BackColor = Color.LightCoral;

				if (dgv.Rows[i].Cells["dgv_HBT_cln_JobContent"].Value != null && !String.IsNullOrEmpty(dgv.Rows[i].Cells["dgv_HBT_cln_JobContent"].Value.ToString()))
					dgv.Rows[i].Cells["dgv_HBT_cln_JobContent"].Style.BackColor = Color.White;
				else
					dgv.Rows[i].Cells["dgv_HBT_cln_JobContent"].Style.BackColor = Color.LightCoral;

				if (dgv.Rows[i].Cells["dgv_HBT_cln_StartMonth"].Value != null && !String.IsNullOrEmpty(dgv.Rows[i].Cells["dgv_HBT_cln_StartMonth"].Value.ToString()))
					dgv.Rows[i].Cells["dgv_HBT_cln_StartMonth"].Style.BackColor = Color.White;
				else
					dgv.Rows[i].Cells["dgv_HBT_cln_StartMonth"].Style.BackColor = Color.LightCoral;

				if (dgv.Rows[i].Cells["dgv_HBT_cln_Salary"].Value != null && !String.IsNullOrEmpty(dgv.Rows[i].Cells["dgv_HBT_cln_Salary"].Value.ToString()))
					dgv.Rows[i].Cells["dgv_HBT_cln_Salary"].Style.BackColor = Color.White;
				else
					dgv.Rows[i].Cells["dgv_HBT_cln_Salary"].Style.BackColor = Color.LightCoral;
			}
		}


		/// <summary>
		/// 計算每月攤提至List
		/// </summary>
		/// <param name="list">裝有ANBTC資料的List</param>
		/// <returns></returns>
		public static List<decimal[]> GetMonthlyData(List<Object> list)
        {
            List<decimal[]> resultlist = new List<decimal[]>();
            decimal[] monthlyData1 = new decimal[13];   // 直接人員每月增減薪資
			decimal[] monthlyData2 = new decimal[13];   // 間接人員每月增減薪資
			decimal[] monthlyData3 = new decimal[13];   // 直接人員伙食人數
			decimal[] monthlyData4 = new decimal[13];   // 間接人員伙食人數
			HumanBudget myObject = null;
			decimal SumActNum_dir = 0; // 實際人數總合
			decimal SumActNum_indir = 0; // 實際人數總合


			for (int j = 0; j < list.Count; j++)
			{
				myObject = (HumanBudget)list[j];

				if ("N".Equals(myObject.IsDelete)) // 刪除標記沒有勾選的才計算
				{
					if ("直接人員".Equals(myObject.Role))
						SumActNum_dir += myObject.ActNum;

					else if ("間接人員".Equals(myObject.Role))
						SumActNum_indir += myObject.ActNum;

					for (int i = 0; i < 12; i++)
					{
						if (myObject.StartMonth <= (i + 1))
						{
							if ("直接人員".Equals(myObject.Role))
							{
								monthlyData1[i] += myObject.TotalChangeSalary;
								//monthlyData1[12] += monthlyData1[i];
								monthlyData1[12] += myObject.TotalChangeSalary;
								//monthlyData1[12] += monthlyData1[i];
							}
							else if ("間接人員".Equals(myObject.Role))
							{
								monthlyData2[i] += myObject.TotalChangeSalary;
								//monthlyData2[12] += monthlyData2[i];
								monthlyData2[12] += myObject.TotalChangeSalary;								
								//monthlyData2[12] += monthlyData2[i];
							}
						}
					}					
				}
            }			

			for (int i = 0; i < 12; i++)
			{
				//myObject = (HumanBudget)list[i];
				monthlyData3[i] = monthlyData3[i] + SumActNum_dir;
				monthlyData4[i] = monthlyData4[i] + SumActNum_indir;			
				monthlyData3[12] += monthlyData3[i];
				monthlyData4[12] += monthlyData4[i];
			}


			for (int i = 0; i < list.Count; i++) 
			{
				myObject = (HumanBudget)list[i];
				//monthlyData3[i] = monthlyData3[i] + SumActNum_dir;
				//monthlyData4[i] = monthlyData4[i] + SumActNum_indir;
				if ("N".Equals(myObject.IsDelete)) { 
					for (int j = 0; j < 12; j++)
					{

						if (myObject.StartMonth <= (j + 1))
						{
							if ("直接人員".Equals(myObject.Role))
							{
								monthlyData3[j] += myObject.DiffNum;
								monthlyData3[12] += myObject.DiffNum; 
							}

							else if ("間接人員".Equals(myObject.Role)) { 
								monthlyData4[j] += myObject.DiffNum;
								monthlyData4[12] += myObject.DiffNum;
							}
						}
					}
				}				
			}			

			resultlist.Add(monthlyData1);     // 直接人員
            resultlist.Add(monthlyData2);     // 間接人員
			resultlist.Add(monthlyData3);     // 直接人員
			resultlist.Add(monthlyData4);     // 間接人員

			return resultlist;
        }

		/// <summary>
		/// 從DB載入ANBTC至List
		/// </summary>
		/// <param name="annualBudgetFormID">人力需求預算表的代碼</param>
		/// <param name="deptNo">編寫年度預算的目標部門代號</param>
		/// <param name="year">年份</param>
		/// <returns></returns>
		public static List<Object> LoadData(string annualBudgetFormID, string deptNo, string year) 
		{
		
			List<Object> list = null;
			StringBuilder SQL = new StringBuilder();
			SqlDataAdapter adapter;
			DataTable dt = new DataTable();
			EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
			SqlConnectionStringBuilder sqlsb;

			sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["AB_ConnString"].ConnectionString);

			sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

			using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
			{
				SqlCommand command = null;
				try
				{					
					SQL.Clear();
					SQL.AppendFormat("SELECT TC000, TC001, TC002, TC003, TC004, " +
							    	 "TC005, TC006, TC007, TC008, TC009, " +
									 "TC010, TC011, TC012, TC013, TC014, " +
									 "TC015, TC016 " +
									 "FROM dbo.ANBTC " +
									 "WHERE TC001 = '{0}' AND TC002 = '{1}' " +
									 "AND TC003 = '{2}' AND TC015 = 'N' ", annualBudgetFormID, deptNo, year);

					if (conn.State != ConnectionState.Open)
						conn.Open();
					command = new SqlCommand(SQL.ToString(), conn);
										
					adapter = new SqlDataAdapter(command);

					adapter.Fill(dt);

					if (dt.Rows.Count > 0) {

						list = new List<object>();
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							ANBTC anbtc = new ANBTC();
							anbtc.Tc000 = Convert.ToDecimal(dt.Rows[i]["TC000"]);
							anbtc.Tc001 = dt.Rows[i]["TC001"].ToString();
							anbtc.Tc002 = dt.Rows[i]["TC002"].ToString();
							anbtc.Tc003 = dt.Rows[i]["TC003"].ToString();
							anbtc.Tc004 = dt.Rows[i]["TC004"].ToString();
							anbtc.Tc005 = dt.Rows[i]["TC005"].ToString();
							anbtc.Tc006 = dt.Rows[i]["TC006"].ToString();
							anbtc.Tc007 = dt.Rows[i]["TC007"].ToString();
							anbtc.Tc008 = Convert.ToDecimal(dt.Rows[i]["TC008"]);
							anbtc.Tc009 = Convert.ToDecimal(dt.Rows[i]["TC009"]);
							anbtc.Tc010 = Convert.ToDecimal(dt.Rows[i]["TC010"]);
							anbtc.Tc011 = Convert.ToDecimal(dt.Rows[i]["TC011"]);
							anbtc.Tc012 = Convert.ToDecimal(dt.Rows[i]["TC012"]);
							anbtc.Tc013 = Convert.ToDecimal(dt.Rows[i]["TC013"]);
							anbtc.Tc014 = dt.Rows[i]["TC014"].ToString();
							anbtc.Tc015 = dt.Rows[i]["TC015"].ToString();
							anbtc.Tc016 = dt.Rows[i]["TC016"].ToString();


							list.Add(anbtc);
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("載入「ANBTC」時發生錯誤：" + e.Message);					
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return list;
			}
        }

		/// <summary>
		/// 將從ANBTC載入的資料帶入DataGridView裡面
		/// </summary>
		/// <param name="dgv">欲接受資料的DataGridView</param>
		/// <param name="list">帶著ANBTC資料的List</param>
		/// <returns></returns>
		public static DataGridView SetDGV(DataGridView dgv, List<Object> list) {
			DataGridViewComboBoxCell cbc = null;

			try
			{
				dgv.Rows.Clear();

				int listCount = list.Count;
				int rowCount = dgv.Rows.Count;

				// 若DataGridView現行的列(Row)數 <= 即將帶入DataGridView的列數，先行新增DataGridView的列數
				if (rowCount <= listCount)
				{
					for (int i = 0; i < (listCount - rowCount) + 1; i++)					
						dgv.Rows.Add();					
				}


				if (list != null && list.Count > 0)
				{
					for (int i = 0; i < list.Count; i++)
					{
						ANBTC anbtc = (ANBTC)list[i];

						cbc = (DataGridViewComboBoxCell)dgv.Rows[i].Cells["dgv_HBT_cln_Role"];

						for (int j = 0; j < cbc.Items.Count; j++)
							if (cbc.Items[j].ToString().Equals(anbtc.Tc004))
								cbc.Value = cbc.Items[j];   // 直接間接人員

						//dgv.Rows[i].Cells[0].Value = anbtc.Tc004;
						dgv.Rows[i].Cells["dgv_HBT_cln_JobTitle"].Value = anbtc.Tc005;		// 職稱
						dgv.Rows[i].Cells["dgv_HBT_cln_Rank"].Value = anbtc.Tc006;			// 職等
						dgv.Rows[i].Cells["dgv_HBT_cln_JobContent"].Value = anbtc.Tc007;    // 工作職掌
						dgv.Rows[i].Cells["dgv_HBT_cln_ActNum"].Value = anbtc.Tc008;		// 實際人數
						dgv.Rows[i].Cells["dgv_HBT_cln_EstNum"].Value = anbtc.Tc009;		// 計畫人數					

						cbc = (DataGridViewComboBoxCell)dgv.Rows[i].Cells["dgv_HBT_cln_StartMonth"];
						for (int j = 0; j < cbc.Items.Count; j++)
							if (cbc.Items[j].ToString().Equals(anbtc.Tc010.ToString()))
								cbc.Value = cbc.Items[j];   // 增減起始月份

						dgv.Rows[i].Cells["dgv_HBT_cln_DiffNum"].Value = anbtc.Tc011;			// 增減人數
						dgv.Rows[i].Cells["dgv_HBT_cln_Salary"].Value = anbtc.Tc012;			// 增減之人員每月薪資
						dgv.Rows[i].Cells["dgv_HBT_cln_TotalChangeSalary"].Value = anbtc.Tc013; // 每月增減薪資
						dgv.Rows[i].Cells["dgv_HBT_cln_Reason"].Value = anbtc.Tc014;            // 增減人力原因	
						dgv.Rows[i].Cells["dgv_HBT_cln_Ref_ANBTN"].Value = anbtc.Tc016;         // 參照號碼
						dgv.Rows[i].Cells["dgv_HBT_cln_Id"].Value = anbtc.Tc000;				// 唯一識別序號
					}
				}
				ANBTC_Model.CheckDataValue(dgv);
			}
			catch (Exception ex)
			{
				_log.Debug("ANBTC_Model的SetDGV方式發生錯誤：" + ex.Message);
				MessageBox.Show("ANBTC_Model的SetDGV方式發生錯誤：" + ex.Message);
			}

			return dgv;			
		}


		/// <summary>
		/// 變更(表三-人力需求表)中，由(表二-組織編制表)帶過去的資料欄位底色及唯讀屬性
		/// </summary>
		/// <param name="dgv"></param>
		public static void SetSyncRowColor(DataGridView dgv) {
			if (dgv != null && dgv.RowCount > 0) {
				for (int i = 0; i < dgv.RowCount; i++) {
					if (dgv.Rows[i].Cells["dgv_HBT_cln_Ref_ANBTN"].Value != null) {
						if (!String.IsNullOrEmpty(dgv.Rows[i].Cells["dgv_HBT_cln_Ref_ANBTN"].Value.ToString())) {
 
							dgv.Rows[i].Cells["dgv_HBT_cln_JobTitle"].ReadOnly = true;  // 職稱
							dgv.Rows[i].Cells["dgv_HBT_cln_JobTitle"].Style.BackColor = Color.PaleGreen;
  
							dgv.Rows[i].Cells["dgv_HBT_cln_Rank"].ReadOnly = true; // 職等
							dgv.Rows[i].Cells["dgv_HBT_cln_Rank"].Style.BackColor = Color.PaleGreen;
 
							dgv.Rows[i].Cells["dgv_HBT_cln_ActNum"].ReadOnly = true;   // 實際人數
							dgv.Rows[i].Cells["dgv_HBT_cln_ActNum"].Style.BackColor = Color.PaleGreen;
  
							dgv.Rows[i].Cells["dgv_HBT_cln_EstNum"].ReadOnly = true;   // 計畫人數
							dgv.Rows[i].Cells["dgv_HBT_cln_EstNum"].Style.BackColor = Color.PaleGreen;
 
							dgv.Rows[i].Cells["dgv_HBT_cln_DiffNum"].ReadOnly = true;  // 增減人數
							dgv.Rows[i].Cells["dgv_HBT_cln_DiffNum"].Style.BackColor = Color.PaleGreen;

							dgv.Rows[i].Cells["dgv_HBT_cln_StartMonth"].ReadOnly = true;	// 起始月份
														 
							dgv.Rows[i].Cells["dgv_HBT_cln_Reason"].ReadOnly = true;    // 增減人力原因
							dgv.Rows[i].Cells["dgv_HBT_cln_Reason"].Style.BackColor = Color.PaleGreen;

							dgv.Rows[i].Cells["dgv_HBT_cbc_Delete"].ReadOnly = true;        // 刪除標記

						}
					}
				}
			}
		}
    }
}
