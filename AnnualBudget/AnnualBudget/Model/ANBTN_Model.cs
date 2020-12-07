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
    class ANBTN_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// 將人力需求預算表DataGridView的資料儲存至List
		/// </summary>
		/// <param name="dgv">組織編制表的DataGridView</param>
		/// <param name="AnnualBudgetFormID">組織編制表的代碼</param>
		/// <param name="mainDeptNo">編寫年度預算的目標部門代號</param>
		/// <param name="deptNo">登入者的部門代號</param>
		/// <param name="year">年份</param>
		/// <param name="userId">登入者的ID</param>
		/// <param name="modidate">修改日期</param>
		/// <returns></returns>
		public static List<Object> SaveMatrixData(DataGridView dgv, string AnnualBudgetFormID, string mainDeptNo, string deptNo, string year, string userId, string modidate)
		{
			List<Object> list = new List<Object>();
			ANBTN myObject = null;

			try
			{

				for (int i = 0; i < dgv.RowCount - 1; i++)
				{
					myObject = new ANBTN();

					myObject.Creator = userId;
					myObject.Usr_group = deptNo;

					myObject.Modifier = userId;
					myObject.Modi_date = modidate;

					myObject.Tn002 = mainDeptNo;
					myObject.Tn003 = year;

					myObject.Tn001 = AnnualBudgetFormID;

					//myObject.No = Convert.ToDecimal(dgv.Rows[i].HeaderCell.Value);

					if (dgv.Rows[i].Cells["dgv_Org_cln_Dept"].Value != null)
						myObject.Tn004 = dgv.Rows[i].Cells["dgv_Org_cln_Dept"].Value.ToString();

					if (dgv.Rows[i].Cells["dgv_Org_cln_JobTitle"].Value != null)
						myObject.Tn005 = dgv.Rows[i].Cells["dgv_Org_cln_JobTitle"].Value.ToString();

					if (dgv.Rows[i].Cells["dgv_Org_cln_Rank"].Value != null)
						myObject.Tn006 = dgv.Rows[i].Cells["dgv_Org_cln_Rank"].Value.ToString();

					
					myObject.Tn007 = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_Org_cln_ActNum"].Value);
					if (dgv.Rows[i].Cells["dgv_Org_cln_OnJob"].Value != null)
						myObject.Tn008 = dgv.Rows[i].Cells["dgv_Org_cln_OnJob"].Value.ToString();
					myObject.Tn009 = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_Org_cln_EstNum"].Value);

					
					myObject.Tn010 = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_Org_cln_DiffNum"].Value);
					myObject.Tn011 = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_Org_cbc_DiffMonth"].Value);
					if (dgv.Rows[i].Cells["dgv_Org_cln_Reason"].Value != null)
						myObject.Tn012 = dgv.Rows[i].Cells["dgv_Org_cln_Reason"].Value.ToString();
					

					DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dgv.Rows[i].Cells["dgv_Org_cbc_Delete"];
					if (chk.Value != null && (bool)chk.Value)
						myObject.Tn013 = "Y";
					else
						myObject.Tn013 = "N";

					if (dgv.Rows[i].Cells["dgv_Org_cln_Ref_ANBTN"].Value != null)
						myObject.Tn014 = dgv.Rows[i].Cells["dgv_Org_cln_Ref_ANBTN"].Value.ToString();
					

					myObject.Tn000 = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_Org_cln_Id"].Value);

					list.Add(myObject);
				}
			}

			catch (Exception e)
			{
				_log.Debug("儲存[組織編制表]時發生錯誤：" + e.Message);
				MessageBox.Show("儲存[組織編制表]時發生錯誤：" + e.Message);
			}


			return list;
		}


		public static bool UpdateToANBTN(List<Object> list)
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
					foreach (ANBTN myObj in list)
					{
						SQL_1.Clear();
						SQL_2.Clear();

						SQL_1.AppendFormat("IF EXISTS (SELECT TN000 FROM dbo.ANBTN WHERE TN000 = {0}) " +
											   "UPDATE dbo.ANBTN SET FLAG = FLAG + 1, MODIFIER = '{1}', MODI_DATE = '{2}', " +
											   "TN004 = '{3}', TN005 = '{4}', TN006 = '{5}', TN007 = {6}, TN008 = '{7}', " +
											   "TN009 = {8}, TN010 = {9}, TN011 = {10}, TN012 = '{11}', TN013 = '{12}', " +
											   "TN014 = '{13}' " +
											   "WHERE TN000 = {14} ELSE ",
											   myObj.Tn000, myObj.Modifier, myObj.Modi_date,
											   myObj.Tn004, myObj.Tn005, myObj.Tn006, myObj.Tn007, myObj.Tn008,
											   myObj.Tn009, myObj.Tn010, myObj.Tn011, myObj.Tn012, myObj.Tn013,
											   myObj.Tn014, 
											   myObj.Tn000);


						SQL_2.AppendFormat("INSERT INTO dbo.ANBTN " +
							"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, " +
							"TN001, TN002, TN003, TN004, TN005, TN006, " +
							"TN007, TN008, TN009, TN010, TN011, TN012, " +
							"TN013, TN014 ) " +
							"VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, " +
							"'{7}', '{8}', '{9}', '{10}', '{11}', " +
							"'{12}', {13}, '{14}', {15}, {16}, " +
							"{17}, '{18}', '{19}', '{20}' ) ",
							myObj.Company, myObj.Creator, myObj.Usr_group, myObj.Create_date, "", "", myObj.Flag,
							myObj.Tn001, myObj.Tn002, myObj.Tn003, myObj.Tn004, myObj.Tn005, 
							myObj.Tn006, myObj.Tn007, myObj.Tn008, myObj.Tn009, myObj.Tn010, 
							myObj.Tn011, myObj.Tn012, myObj.Tn013, myObj.Tn014);

						SQL_1 = SQL_1.Append(SQL_2);


						//command = new SqlCommand(SQL_1.ToString(), conn);
						command.CommandText = SQL_1.ToString();
						command.ExecuteNonQuery();
					}
					trans.Commit();

				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTN」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「ANBTN」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTN」時發生錯誤的訊息：" + e2.Message);
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
		/// 從DB載入ANBTN至List
		/// </summary>
		/// <param name="annualBudgetFormID">組織編制表的代碼</param>
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
					SQL.AppendFormat("SELECT TN000, TN001, TN002, TN003, TN004, " +
									 "TN005, TN006, TN007, TN008, TN009, " +
									 "TN010, TN011, TN012, TN013, TN014 " +									 
									 "FROM dbo.ANBTN " +
									 "WHERE TN001 = '{0}' AND TN002 = '{1}' " +
									 "AND TN003 = '{2}' AND TN013 = 'N' ", annualBudgetFormID, deptNo, year);

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
							ANBTN anbtn = new ANBTN();
							anbtn.Tn000 = Convert.ToDecimal(dt.Rows[i]["TN000"]);
							anbtn.Tn001 = dt.Rows[i]["TN001"].ToString();
							anbtn.Tn002 = dt.Rows[i]["TN002"].ToString();
							anbtn.Tn003 = dt.Rows[i]["TN003"].ToString();
							anbtn.Tn004 = dt.Rows[i]["TN004"].ToString();
							anbtn.Tn005 = dt.Rows[i]["TN005"].ToString();
							anbtn.Tn006 = dt.Rows[i]["TN006"].ToString();
							anbtn.Tn007 = Convert.ToDecimal(dt.Rows[i]["TN007"].ToString());
							anbtn.Tn008 = dt.Rows[i]["TN008"].ToString();
							anbtn.Tn009 = Convert.ToDecimal(dt.Rows[i]["TN009"]);
							anbtn.Tn010 = Convert.ToDecimal(dt.Rows[i]["TN010"]);
							anbtn.Tn011 = Convert.ToDecimal(dt.Rows[i]["TN011"]);
							anbtn.Tn012 = dt.Rows[i]["TN012"].ToString();
							anbtn.Tn013 = dt.Rows[i]["TN013"].ToString();
							anbtn.Tn014 = dt.Rows[i]["TN014"].ToString();
							
							list.Add(anbtn);
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("載入「ANBTN」時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return list;
			}
		}


		/// <summary>
		/// 將從ANBTN載入的資料塞入DataGridView裡面
		/// </summary>
		/// <param name="dgv">欲接受資料的DataGridView</param>
		/// <param name="list">帶著ANBTN資料的List</param>
		/// <returns></returns>
		public static DataGridView SetDGV(DataGridView dgv, List<Object> list)
		{
			DataGridViewComboBoxCell cbc = null;

			try
			{
				dgv.Rows.Clear();

				int listCount = list.Count;
				int rowCount = dgv.Rows.Count;

				if (rowCount <= listCount)
				{
					for (int i = 0; i < (listCount - rowCount) + 1; i++)
						dgv.Rows.Add();
				}


				if (list != null && list.Count > 0)
				{
					for (int i = 0; i < list.Count; i++)
					{
						ANBTN obj = (ANBTN)list[i];

						dgv.Rows[i].Cells["dgv_Org_cln_Dept"].Value = obj.Tn004;	// 部門/單位
						dgv.Rows[i].Cells["dgv_Org_cln_JobTitle"].Value = obj.Tn005;   // 職稱職位
						dgv.Rows[i].Cells["dgv_Org_cln_Rank"].Value = obj.Tn006;   // 職等職級
						dgv.Rows[i].Cells["dgv_Org_cln_ActNum"].Value = obj.Tn007;   // 實際人數
						dgv.Rows[i].Cells["dgv_Org_cln_OnJob"].Value = obj.Tn008;   // 現職者
						dgv.Rows[i].Cells["dgv_Org_cln_EstNum"].Value = obj.Tn009;   // 計畫人數					
						dgv.Rows[i].Cells["dgv_Org_cln_DiffNum"].Value = obj.Tn010;   // 計畫與實際差額					

						// 增減月份
						cbc = (DataGridViewComboBoxCell)dgv.Rows[i].Cells["dgv_Org_cbc_DiffMonth"];
						for (int j = 0; j < cbc.Items.Count; j++)
							if (cbc.Items[j].ToString().Equals(obj.Tn011.ToString()))
								cbc.Value = cbc.Items[j];   

						dgv.Rows[i].Cells["dgv_Org_cln_Reason"].Value = obj.Tn012;  // 增減人力原因
						dgv.Rows[i].Cells["dgv_Org_cln_Ref_ANBTN"].Value = obj.Tn014;  // 參照號碼
						dgv.Rows[i].Cells["dgv_Org_cln_Id"].Value = obj.Tn000;  // 唯一識別序號
					}
				}
			}
			catch (Exception ex)
			{
				_log.Debug("ANBTN_Model的SetDGV方式發生錯誤：" + ex.Message);
				MessageBox.Show("ANBTN_Model的SetDGV方式發生錯誤：" + ex.Message);
			}

			return dgv;
		}


		public static string Get_Ref_TNID()
		{
			string result = "";

			int num = 0;
			StringBuilder selectString = new StringBuilder();
			SqlCommand command = null;
			SqlDataReader reader = null;
			string dateTime = DateTime.Now.ToString("yyMM");
			EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
			SqlConnectionStringBuilder sqlsb;

			sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["AB_ConnString"].ConnectionString);

			sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

			using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
			{
				try
				{
					selectString.Clear();
					selectString.AppendFormat("SELECT TOP 1 TN014 FROM dbo.ANBTN WHERE SUBSTRING(TN014, 0, 5) = '{0}' GROUP BY TN014 ORDER BY TN014 desc", dateTime);

					conn.Open();
					command = new SqlCommand(selectString.ToString(), conn);
					reader = command.ExecuteReader();

					if (reader.HasRows)
					{
						while (reader.Read())
							result = reader.GetString(0);

						if (!String.IsNullOrEmpty(result))
						{
							num = Convert.ToInt32(result);
							num += 1;
							result = num.ToString();
						}
						else
							result = dateTime + "0001";

					}
					else
						result = dateTime + "0001";
				}
				catch (Exception e)
				{
					_log.Debug("執行[Get_Ref_TNID]時發生錯誤：" + e.Message);
					MessageBox.Show("執行[Get_Ref_TNID]時發生錯誤：" + e.Message);
				}
				finally
				{
					if (reader != null) reader.Close();
					if (command != null) command.Dispose();
				}
			}
			return result;
		}


		public static void SyncDataToDGV_HBT(List<Object> list, DataGridView dgv) 
		{
			bool isMatch = false;
			for (int i = 0; i < list.Count; i++) 
			{
				ANBTN anbtn = (ANBTN)list[i];
				isMatch = false;
				for (int j = 0; j < dgv.RowCount - 1; j++) {
					string RefCln = "";
					if (dgv.Rows[j].Cells["dgv_HBT_cln_Ref_ANBTN"].Value != null) { 
						RefCln = dgv.Rows[j].Cells["dgv_HBT_cln_Ref_ANBTN"].Value.ToString();

						if (anbtn.Tn014.Equals(RefCln))
						{
							isMatch = true;

							// 把資料對應貼過來
							SetValue(dgv, j, anbtn);							

							break;
						}						
					}
				}

				if (isMatch == false)   // 沒有對應到的資料
				{					
					if (!"Y".Equals(anbtn.Tn013))	// 且標記不為刪除
					{
						// 把資料新增過來
						dgv.Rows.Add();
						SetValue(dgv, dgv.RowCount - 2, anbtn);
					}
				}
			}
		}

		private static void SetValue(DataGridView dgv, int ClnNum, ANBTN anbtn)
		{
			dgv.Rows[ClnNum].Cells["dgv_HBT_cln_JobTitle"].Value = anbtn.Tn005;  // 職稱
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_JobTitle"].ReadOnly = true;
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_JobTitle"].Style.BackColor = Color.PaleGreen;
			dgv.Rows[ClnNum].Cells["dgv_HBT_cln_Rank"].Value = anbtn.Tn006;      // 職等
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_Rank"].ReadOnly = true;
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_Rank"].Style.BackColor = Color.PaleGreen;
			dgv.Rows[ClnNum].Cells["dgv_HBT_cln_ActNum"].Value = anbtn.Tn007;    // 實際人數
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_ActNum"].ReadOnly = true;
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_ActNum"].Style.BackColor = Color.PaleGreen;
			dgv.Rows[ClnNum].Cells["dgv_HBT_cln_EstNum"].Value = anbtn.Tn009;    // 計畫人數
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_EstNum"].ReadOnly = true;
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_EstNum"].Style.BackColor = Color.PaleGreen;
			dgv.Rows[ClnNum].Cells["dgv_HBT_cln_DiffNum"].Value = anbtn.Tn009 - anbtn.Tn007;    // 增減人數
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_DiffNum"].ReadOnly = true;
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_DiffNum"].Style.BackColor = Color.PaleGreen;

			// 增減起始月份
			DataGridViewComboBoxCell cbc = (DataGridViewComboBoxCell)dgv.Rows[ClnNum].Cells["dgv_HBT_cln_StartMonth"];
			/*
			if (cbc.Value != null && !String.IsNullOrEmpty(cbc.Value.ToString()))
			{
				for (int i = 0; i < cbc.Items.Count; i++)
				{
					if (cbc.Items[i].ToString().Equals(anbtn.Tn011.ToString()))
						cbc.Value = cbc.Items[i];
				}
			}*/
			if (cbc.Items.Count >=0 )
			{
				for (int i = 0; i < cbc.Items.Count; i++)
				{
					if (cbc.Items[i].ToString().Equals(anbtn.Tn011.ToString()))
						cbc.Value = cbc.Items[i];
				}
			}



			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_StartMonth"].ReadOnly = true;

			dgv.Rows[ClnNum].Cells["dgv_HBT_cln_Reason"].Value = anbtn.Tn012;    // 增減人力原因
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_Reason"].ReadOnly = true;
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cln_Reason"].Style.BackColor = Color.PaleGreen;
			dgv.Rows[ClnNum].Cells["dgv_HBT_cln_Ref_ANBTN"].Value = anbtn.Tn014;    // 參照號碼

			DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dgv.Rows[ClnNum].Cells["dgv_HBT_cbc_Delete"];
			if ("Y".Equals(anbtn.Tn013)) { 
				chk.Value = true;
				dgv.Rows[ClnNum].Visible = false;	// 若帶過來的資料已刪除，則這筆資料隱藏
			}
			else
				chk.Value = false;
			//dgv.Rows[ClnNum].Cells["dgv_HBT_cbc_Delete"].ReadOnly = true;        // 刪除標記
		}
	}
}