using AnnualBudget.BOs;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AnnualBudget.Model
{
    class ANBTD_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// 新增或更新資料至ANBTD資料表
		/// </summary>
		/// <param name="list">存放ANBTD物件的List</param>
		/// <returns></returns>
		public static bool UpdateToANBTD(List<Object> myList)
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
					foreach (EmployeeTraining myObj in myList) { 
						SQL_1.Clear();
						SQL_2.Clear();

						SQL_1.AppendFormat("IF EXISTS (SELECT TD000 FROM dbo.ANBTD WHERE TD000 = {0} ) " +
												"UPDATE dbo.ANBTD SET FLAG = FLAG + 1, MODIFIER = '{1}', MODI_DATE = '{2}', " +
												"TD004 = '{3}', TD005 = '{4}', TD006 = '{5}', TD007 = '{6}', TD008 = '{7}', " +
												"TD009 = {8}, TD010 = {9}, TD011 = {10}, TD012 = {11}, TD013 = {12}, " +
												"TD014 = '{13}', TD015 = '{14}' " +
												"WHERE TD000 = {15} ELSE ",
												myObj.Id, myObj.Modifier, myObj.Modi_date, 
												myObj.Ability, myObj.FocalPoint, myObj.Target, myObj.PeopleNum, myObj.In_or_Ex,
												myObj.StartMonth, myObj.EndMonth, myObj.Hours, myObj.MonthlyFee, myObj.ProjectQuota, 
												myObj.Memo, myObj.IsDelete, myObj.Id);

						SQL_2.AppendFormat("INSERT INTO dbo.ANBTD " +
							"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, " +
							"TD001, TD002, TD003, TD004, TD005, TD006, " +
							"TD007, TD008, TD009, TD010, TD011, TD012, " +
							"TD013, TD014, TD015 ) " +
							"VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, " +
							"'{7}', '{8}', '{9}', '{10}', '{11}', " +
							"'{12}', {13}, '{14}', {15}, {16}, " +
							"{17}, {18}, {19}, '{20}', '{21}' ) ",
							myObj.Company, myObj.Creator, myObj.Usr_group, myObj.Create_date, "", "", myObj.Flag,
							myObj.AnnualBudgetFormID, myObj.DeptNo, myObj.Year, myObj.Ability, myObj.FocalPoint, myObj.Target, 
							myObj.PeopleNum, myObj.In_or_Ex, myObj.StartMonth, myObj.EndMonth, myObj.Hours, myObj.MonthlyFee,
							myObj.ProjectQuota, myObj.Memo, myObj.IsDelete);

						SQL_1 = SQL_1.Append(SQL_2);

						command.CommandText = SQL_1.ToString();
						command.ExecuteNonQuery();
						
					}
					trans.Commit();
				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTD」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2) {
						_log.Debug("更新至「ANBTD」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTD」時發生錯誤的訊息：" + e2.Message);
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
		/// 將教育訓練計劃表DataGridView的資料儲存至List
		/// </summary>
		/// <param name="dgv">教育訓練計劃表的DataGridView</param>
		/// <param name="AnnualBudgetFormID">教育訓練計劃表的代碼</param>
		/// <param name="mainDeptNo">編寫年度預算的目標部門代號</param>
		/// <param name="deptNo">登入者的部門代號</param>
		/// <param name="year">年份</param>
		/// <param name="userId">登入者的ID</param>
		/// <param name="modidate">修改日期</param>
		/// <returns></returns>
		public static List<Object> SaveMatrixData(DataGridView dgv, string AnnualBudgetFormID, string MainDeptNo, string deptNo, string year, string userId, string modidate)
		{
			List<Object> list = new List<Object>();
			EmployeeTraining myObject = null;
            try { 
				for (int i = 0; i < dgv.RowCount - 1; i++)
				{
					myObject = new EmployeeTraining();
					myObject.Creator = userId;
					myObject.Usr_group = deptNo;
					myObject.Modifier = userId;
					myObject.Modi_date = modidate;
					myObject.DeptNo = MainDeptNo;
					myObject.Year = year;
					myObject.AnnualBudgetFormID = AnnualBudgetFormID;

					myObject.No = Convert.ToDecimal(dgv.Rows[i].HeaderCell.Value);

					if (dgv.Rows[i].Cells[0].Value != null)
						myObject.Ability = dgv.Rows[i].Cells[0].Value.ToString();

					if (dgv.Rows[i].Cells[1].Value != null)
						myObject.FocalPoint = dgv.Rows[i].Cells[1].Value.ToString();

					if (dgv.Rows[i].Cells[2].Value != null)
						myObject.Target = dgv.Rows[i].Cells[2].Value.ToString();

					myObject.PeopleNum = Convert.ToDecimal(dgv.Rows[i].Cells[3].Value);

					if (dgv.Rows[i].Cells[4].Value != null)
						myObject.In_or_Ex = dgv.Rows[i].Cells[4].Value.ToString();


					myObject.StartMonth = Convert.ToDecimal(dgv.Rows[i].Cells[5].Value);
					myObject.EndMonth = Convert.ToDecimal(dgv.Rows[i].Cells[6].Value);
					myObject.Hours = Convert.ToDecimal(dgv.Rows[i].Cells[7].Value);
					myObject.MonthlyFee = Convert.ToDecimal(dgv.Rows[i].Cells[8].Value);
					myObject.ProjectQuota = Convert.ToDecimal(dgv.Rows[i].Cells[9].Value);
					if (dgv.Rows[i].Cells[10].Value != null)
						myObject.Memo = dgv.Rows[i].Cells[10].Value.ToString();

					DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dgv.Rows[i].Cells[11];
					if (chk.Value != null && (bool)chk.Value)
						myObject.IsDelete = "Y";
					else
						myObject.IsDelete = "N";

					myObject.Id = Convert.ToDecimal(dgv.Rows[i].Cells[12].Value);

					list.Add(myObject);
				}
			}
			catch (Exception e)
			{
				_log.Debug("儲存[教育訓練計劃表]時發生錯誤：" + e.Message);
				MessageBox.Show("儲存[教育訓練計劃表]時發生錯誤：" + e.Message);
			}

			return list;
        }

		/// <summary>
		/// 計算每月攤提至List
		/// </summary>
		/// <param name="list">存放ANBTD物件的List</param>
		/// <returns></returns>
		public static decimal[] GetMonthlyData(List<Object> list)
        {
            decimal[] monthlyData = new decimal[13];
            EmployeeTraining myObject = null;

			
			for (int j = 0; j < list.Count; j++)    // 資料筆數
			{
				myObject = (EmployeeTraining)list[j];			

				if ("N".Equals(myObject.IsDelete)) // 刪除標記沒有勾選的才計算
				{
					for (int i = 0; i < 12; i++)    // 月份
					{
						if (myObject.StartMonth <= (i + 1) && myObject.EndMonth >= (i + 1))
						{
							monthlyData[i] += myObject.MonthlyFee;
							monthlyData[12] += myObject.MonthlyFee;
						}
					}
				}
            }

            return monthlyData;
        }

		/// <summary>
		/// 從DB載入ANBTD的資料，並帶入至List
		/// </summary>
		/// <param name="annualBudgetFormID">教育訓練計劃表的代碼</param>
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
					SQL.AppendFormat("SELECT TD000, TD001, TD002, TD003, TD004, " +
									 "TD005, TD006, TD007, TD008, TD009, " +
									 "TD010, TD011, TD012, TD013, TD014 " +
									 "FROM dbo.ANBTD " +
									 "WHERE TD001 = '{0}' AND TD002 = '{1}' " +
									 "AND TD003 = '{2}' AND TD015 = 'N' ", annualBudgetFormID, deptNo, year);

					if (conn.State != ConnectionState.Open)
						conn.Open();
					command = new SqlCommand(SQL.ToString(), conn);

					adapter = new SqlDataAdapter(command);

					adapter.Fill(dt);

					if (dt.Rows.Count > 0)
					{

						list = new List<object>();
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							ANBTD myObject = new ANBTD();
							myObject.Td000 = Convert.ToDecimal(dt.Rows[i]["TD000"]);   // 流水號
							myObject.Td001 = dt.Rows[i]["TD001"].ToString();           // 單別代號
							myObject.Td002 = dt.Rows[i]["TD002"].ToString();           // 部門代號
							myObject.Td003 = dt.Rows[i]["TD003"].ToString();           // 年度
							myObject.Td004 = dt.Rows[i]["TD004"].ToString();
							myObject.Td005 = dt.Rows[i]["TD005"].ToString();
							myObject.Td006 = dt.Rows[i]["TD006"].ToString();
							myObject.Td007 = Convert.ToDecimal(dt.Rows[i]["TD007"]);
							myObject.Td008 = dt.Rows[i]["TD008"].ToString();
							myObject.Td009 = Convert.ToDecimal(dt.Rows[i]["TD009"]);
							myObject.Td010 = Convert.ToDecimal(dt.Rows[i]["TD010"]);
							myObject.Td011 = Convert.ToDecimal(dt.Rows[i]["TD011"]);
							myObject.Td012 = Convert.ToDecimal(dt.Rows[i]["TD012"]);
							myObject.Td013 = Convert.ToDecimal(dt.Rows[i]["TD013"]);
							myObject.Td014 = dt.Rows[i]["TD014"].ToString();							


							list.Add(myObject);
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("載入[ANBTD]資料時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return list;
			}
		}


		/// <summary>
		/// 將從ANBTD載入的資料帶入DataGridView裡面
		/// </summary>
		/// <param name="dgv">欲接受資料的DataGridView</param>
		/// <param name="list">存放ANBTD物件的List</param>
		/// <returns></returns>
		public static DataGridView SetDGV(DataGridView dgv, List<Object> list)
		{
			DataGridViewComboBoxCell cbc = null;
			try
			{
				dgv.Rows.Clear();

				if (list != null && list.Count > 0)
				{
					int listCount = list.Count;
					int rowCount = dgv.Rows.Count;

					// 若DataGridView現行的列(Row)數 <= 即將帶入DataGridView的列數，先行新增DataGridView的列數
					if (rowCount <= listCount)
						for (int i = 0; i < (listCount - rowCount) + 1; i++)
							dgv.Rows.Add();


					for (int i = 0; i < list.Count; i++)
					{
						ANBTD myObj = (ANBTD)list[i];

						dgv.Rows[i].Cells[0].Value = myObj.Td004;   // 職務能力
						dgv.Rows[i].Cells[1].Value = myObj.Td005;   // 訓練重點
						dgv.Rows[i].Cells[2].Value = myObj.Td006;   // 對象
						dgv.Rows[i].Cells[3].Value = myObj.Td007;   // 人數
												
						cbc = (DataGridViewComboBoxCell)dgv.Rows[i].Cells[4];
						for (int j = 0; j < cbc.Items.Count; j++)
							if (cbc.Items[j].ToString().Equals(myObj.Td008.ToString()))
								cbc.Value = cbc.Items[j];             // 內外訓
												
						cbc = (DataGridViewComboBoxCell)dgv.Rows[i].Cells[5];
						for (int j = 0; j < cbc.Items.Count; j++)
							if (cbc.Items[j].ToString().Equals(myObj.Td009.ToString()))
								cbc.Value = cbc.Items[j];             // 起始月

						cbc = (DataGridViewComboBoxCell)dgv.Rows[i].Cells[6];
						for (int j = 0; j < cbc.Items.Count; j++)
							if (cbc.Items[j].ToString().Equals(myObj.Td010.ToString()))
								cbc.Value = cbc.Items[j];   // 結束月

						dgv.Rows[i].Cells[7].Value = myObj.Td011;   // 時數
						dgv.Rows[i].Cells[8].Value = myObj.Td012;   // 每月費用
						dgv.Rows[i].Cells[9].Value = myObj.Td013;   // 該計劃費用總額
						dgv.Rows[i].Cells[10].Value = myObj.Td014;  // 備註

						dgv.Rows[i].Cells[12].Value = myObj.Td000;  // 唯一序號
					}
				}
			}
			catch (Exception ex)
			{
				_log.Debug("ANBTD_Model的SetDGV方式發生錯誤：" + ex.Message);
				MessageBox.Show("ANBTD_Model的SetDGV方式發生錯誤：" + ex.Message);
			}

			return dgv;
		}
	}
}
