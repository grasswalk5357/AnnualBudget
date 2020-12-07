using AnnualBudget.BOs;
using NPOI.SS.Formula.Functions;
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
    class ANBTI_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// 新增或更新資料至ANBTI資料表
		/// </summary>
		/// <param name="myList">存放ANBTI物件的List</param>
		/// <returns></returns>
		public static bool UpdateToANBTI(List<Object> list)
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
					foreach (ANBTI myObj in list) { 
						SQL_1.Clear();
						SQL_2.Clear();

						SQL_1.AppendFormat("IF EXISTS (SELECT TI000 FROM dbo.ANBTI WHERE TI000 = {0} ) " +
											"UPDATE dbo.ANBTI SET FLAG = FLAG + 1, MODIFIER = '{1}', MODI_DATE = '{2}', " +
											"TI004 = '{3}', TI005 = '{4}', TI006 = {5}, TI007 = '{6}', TI008 = '{7}', " +
											"TI009 = '{8}', TI010 = {9}, TI011 = {10}, TI012 = {11}, TI013 = {12}, " +
											"TI014 = {13}, TI015 = {14}, TI016 = {15}, TI017 = {16}, TI018 = {17}, " +
											"TI019 = {18}, TI020 = {19}, TI021 = {20}, TI022 = {21} " +
											"WHERE TI000 = {22}  ELSE ",
											myObj.Ti000, myObj.Modifier, myObj.Modi_date,
											myObj.Ti004, myObj.Ti005, myObj.Ti006, myObj.Ti007, myObj.Ti008, 
											myObj.Ti009, myObj.Ti010, myObj.Ti011, myObj.Ti012, myObj.Ti013, 
											myObj.Ti014, myObj.Ti015, myObj.Ti016, myObj.Ti017, myObj.Ti018, 
											myObj.Ti019, myObj.Ti020, myObj.Ti021, myObj.Ti022,
											myObj.Ti000);


						SQL_2.AppendFormat("INSERT INTO dbo.ANBTI " +
							"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, " +
							"TI001, TI002, TI003, TI004, TI005, TI006, " +
							"TI007, TI008, TI009, TI010, TI011, TI012, " +
							"TI013, TI014, TI015, TI016, TI017, TI018, " +
							"TI019, TI020, TI021, TI022 " +
							") " +
							"VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, " +
							"'{7}', '{8}', '{9}', '{10}', '{11}', " +
							"{12}, '{13}', '{14}', '{15}', {16}, " +
							"{17}, {18}, {19}, {20}, {21}, " +
							"{22}, {23}, {24}, {25}, {26}, " +
							"{27}, {28} " +
							" ) ",
							myObj.Company, myObj.Creator, myObj.Usr_group, myObj.Create_date, "", "", myObj.Flag,
							myObj.Ti001, myObj.Ti002, myObj.Ti003, myObj.Ti004, myObj.Ti005,
							myObj.Ti006, myObj.Ti007, myObj.Ti008, myObj.Ti009, myObj.Ti010,
							myObj.Ti011, myObj.Ti012, myObj.Ti013, myObj.Ti014, myObj.Ti015,
							myObj.Ti016, myObj.Ti017, myObj.Ti018, myObj.Ti019, myObj.Ti020,
							myObj.Ti021, myObj.Ti022);

						SQL_1 = SQL_1.Append(SQL_2);

						command.CommandText = SQL_1.ToString();
						command.ExecuteNonQuery();
					}
					trans.Commit();

				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTI」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「ANBTI」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTI」時發生錯誤的訊息：" + e2.Message);
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
		/// 將預算總表的資料儲存至List並回傳
		/// </summary>
		/// <param name="dgv"></param>
		/// <param name="dt"></param>
		/// <param name="annualBudgetFormID"></param>
		/// <param name="deptNo"></param>
		/// <param name="year"></param>
		/// <returns></returns>
		public static List<object> SaveMatrixData(DataGridView dgv, DataTable dt, string annualBudgetFormID, string MainDeptNo, string deptNo, string year, string userId, string modidate) 
		{
			List<object> list = new List<object>();

			try
			{
				for (int i = 0; i < dt.Rows.Count; i++)
				{
					ANBTI myObject = new ANBTI();
					myObject.Creator = userId;
					myObject.Usr_group = deptNo;
					myObject.Modifier = userId;
					myObject.Modi_date = modidate;
					myObject.Ti001 = annualBudgetFormID;  // 單別代號
					myObject.Ti002 = MainDeptNo;          // 部門代號
					myObject.Ti003 = year;                // 年度
					myObject.Ti004 = dt.Rows[i][6].ToString();         // 會科樣版號
					myObject.Ti005 = dt.Rows[i][7].ToString();         // 會科樣版版本號

					myObject.Ti006 = Convert.ToDecimal(dt.Rows[i][0]); // 樣版科目流水號
					myObject.Ti007 = dt.Rows[i][1].ToString();          // 會科編號
					myObject.Ti008 = dt.Rows[i][2].ToString();          // 會科名稱
					myObject.Ti009 = dt.Rows[i][3].ToString();          // 會科名稱說明

					myObject.Ti010 = Convert.ToDecimal(dgv.Rows[i].Cells[0].Value);	   // 1月
					myObject.Ti011 = Convert.ToDecimal(dgv.Rows[i].Cells[1].Value);    // 2月
					myObject.Ti012 = Convert.ToDecimal(dgv.Rows[i].Cells[2].Value);    // 3月
					myObject.Ti013 = Convert.ToDecimal(dgv.Rows[i].Cells[3].Value);    // 4月
					myObject.Ti014 = Convert.ToDecimal(dgv.Rows[i].Cells[4].Value);    // 5月
					myObject.Ti015 = Convert.ToDecimal(dgv.Rows[i].Cells[5].Value);    // 6月
					myObject.Ti016 = Convert.ToDecimal(dgv.Rows[i].Cells[6].Value);    // 7月
					myObject.Ti017 = Convert.ToDecimal(dgv.Rows[i].Cells[7].Value);    // 8月
					myObject.Ti018 = Convert.ToDecimal(dgv.Rows[i].Cells[8].Value);    // 9月
					myObject.Ti019 = Convert.ToDecimal(dgv.Rows[i].Cells[9].Value);    // 10月
					myObject.Ti020 = Convert.ToDecimal(dgv.Rows[i].Cells[10].Value);   // 11月
					myObject.Ti021 = Convert.ToDecimal(dgv.Rows[i].Cells[11].Value);   // 12月
					
					myObject.Ti022 = Convert.ToDecimal(dgv.Rows[i].Cells[12].Value);   // 項目總計

					myObject.Ti000 = Convert.ToDecimal(dgv.Rows[i].Cells[13].Value);   // 唯一序號

					list.Add(myObject);
				}
			}
			catch (Exception e)
			{
				_log.Debug("儲存[預算總表]時發生錯誤：" + e.Message);
				MessageBox.Show("儲存[預算總表]時發生錯誤：" + e.Message);
			}


			return list;
			
		}


		/// <summary>
		/// 將預算總表的資料儲存至List並回傳(for 會科預填表單使用)
		/// </summary>
		/// <param name="dgv"></param>
		/// <param name="dt"></param>
		/// <param name="annualBudgetFormID"></param>
		/// <param name="MainDeptNo"></param>
		/// <param name="deptNo"></param>
		/// <param name="year"></param>
		/// <param name="userId"></param>
		/// <param name="modidate"></param>
		/// <returns></returns>
		public static List<object> SaveMatrixDataForAccounting(DataGridView dgv, DataTable dt, string annualBudgetFormID, string MainDeptNo, string deptNo, string year, string userId, string modidate)
		{
			List<object> list = new List<object>();     // 儲存結果
			string dt_ItemFullName = "";
			string dgv_ItemFullName = "";

			try
			{
				for (int i = 0; i < dt.Rows.Count; i++)
				{
					ANBTI myObject = new ANBTI();
					myObject.Creator = userId;
					myObject.Usr_group = deptNo;
					myObject.Modifier = userId;
					myObject.Modi_date = modidate;
					myObject.Ti001 = annualBudgetFormID;  // 單別代號
					myObject.Ti002 = MainDeptNo;          // 部門代號
					myObject.Ti003 = year;                // 年度
					myObject.Ti004 = dt.Rows[i][6].ToString();         // 會科樣版編號
					myObject.Ti005 = dt.Rows[i][7].ToString();         // 會科樣版版本號

					myObject.Ti006 = Convert.ToDecimal(dt.Rows[i][0]); // 樣版科目流水號
					myObject.Ti007 = dt.Rows[i][1].ToString();         // 會科編號
					myObject.Ti008 = dt.Rows[i][2].ToString();         // 會科名稱
					myObject.Ti009 = dt.Rows[i][3].ToString();         // 會科名稱說明


					//把會科編號、會科名稱、會科名稱說明 給組合起來
					if (String.IsNullOrEmpty(myObject.Ti007))
						dt_ItemFullName = myObject.Ti008 + "-(" + myObject.Ti009 + ")"; 
					else
						dt_ItemFullName = myObject.Ti007 + "-" + myObject.Ti008 + "-(" + myObject.Ti009 + ")";


					for (int j = 0; j < dgv.Rows.Count; j++) {

						dgv_ItemFullName = dgv.Rows[j].HeaderCell.Value.ToString();

						if (dt_ItemFullName.Equals(dgv_ItemFullName)) {
							myObject.Ti010 = Convert.ToDecimal(dgv.Rows[j].Cells[0].Value);    // 1月
							myObject.Ti011 = Convert.ToDecimal(dgv.Rows[j].Cells[1].Value);    // 2月
							myObject.Ti012 = Convert.ToDecimal(dgv.Rows[j].Cells[2].Value);    // 3月
							myObject.Ti013 = Convert.ToDecimal(dgv.Rows[j].Cells[3].Value);    // 4月
							myObject.Ti014 = Convert.ToDecimal(dgv.Rows[j].Cells[4].Value);    // 5月
							myObject.Ti015 = Convert.ToDecimal(dgv.Rows[j].Cells[5].Value);    // 6月
							myObject.Ti016 = Convert.ToDecimal(dgv.Rows[j].Cells[6].Value);    // 7月
							myObject.Ti017 = Convert.ToDecimal(dgv.Rows[j].Cells[7].Value);    // 8月
							myObject.Ti018 = Convert.ToDecimal(dgv.Rows[j].Cells[8].Value);    // 9月
							myObject.Ti019 = Convert.ToDecimal(dgv.Rows[j].Cells[9].Value);    // 10月
							myObject.Ti020 = Convert.ToDecimal(dgv.Rows[j].Cells[10].Value);   // 11月
							myObject.Ti021 = Convert.ToDecimal(dgv.Rows[j].Cells[11].Value);   // 12月

							myObject.Ti022 = Convert.ToDecimal(dgv.Rows[j].Cells[12].Value);   // 項目總計

							myObject.Ti000 = Convert.ToDecimal(dgv.Rows[j].Cells[13].Value);   // 唯一序號
							
							break;
						}
						
								
					}

					list.Add(myObject);
				}
			}
			catch (Exception e)
			{
				_log.Debug("儲存[預算總表]時發生錯誤：" + e.Message);
				MessageBox.Show("儲存[預算總表]時發生錯誤：" + e.Message);
			}


			return list;

		}


		public static List<Object> LoadData(string annualBudgetFormID, string deptNo, string year, string TmplID)
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
					SQL.AppendFormat("SELECT TI000, TI001, TI002, TI003, TI004, " +
											"TI005, TI006, TI007, TI008, TI009, " +
											"TI010, TI011, TI012, TI013, TI014, " +
											"TI015, TI016, TI017, TI018, TI019, " +
											"TI020, TI021, TI022 " +
									 "FROM dbo.ANBTI " +
									 "WHERE TI001 = '{0}' AND TI002 = '{1}' " +
									 "AND TI003 = '{2}' AND TI004 = '{3}' ", annualBudgetFormID, deptNo, year, TmplID);

					if (conn.State != ConnectionState.Open)
						conn.Open();
					command = new SqlCommand(SQL.ToString(), conn);

					adapter = new SqlDataAdapter(command);

					adapter.Fill(dt);
					
					if (dt.Rows.Count > 0)
					{

						list = new List<Object>();
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							ANBTI myObject = new ANBTI();
							myObject.Ti000 = Convert.ToDecimal(dt.Rows[i]["Ti000"]);   // 唯一識別流水號
							myObject.Ti001 = dt.Rows[i]["Ti001"].ToString();           // 單別代號
							myObject.Ti002 = dt.Rows[i]["Ti002"].ToString();           // 部門代號
							myObject.Ti003 = dt.Rows[i]["Ti003"].ToString();           // 年度
							myObject.Ti004 = dt.Rows[i]["Ti004"].ToString();
							myObject.Ti005 = dt.Rows[i]["Ti005"].ToString();
							myObject.Ti006 = Convert.ToDecimal(dt.Rows[i]["Ti006"]);
							myObject.Ti007 = dt.Rows[i]["Ti007"].ToString();
							myObject.Ti008 = dt.Rows[i]["Ti008"].ToString();
							myObject.Ti009 = dt.Rows[i]["Ti009"].ToString();
							myObject.Ti010 = Convert.ToDecimal(dt.Rows[i]["Ti010"]);
							myObject.Ti011 = Convert.ToDecimal(dt.Rows[i]["Ti011"]);
							myObject.Ti012 = Convert.ToDecimal(dt.Rows[i]["Ti012"]);
							myObject.Ti013 = Convert.ToDecimal(dt.Rows[i]["Ti013"]);
							myObject.Ti014 = Convert.ToDecimal(dt.Rows[i]["Ti014"]);
							myObject.Ti015 = Convert.ToDecimal(dt.Rows[i]["Ti015"]);
							myObject.Ti016 = Convert.ToDecimal(dt.Rows[i]["Ti016"]);
							myObject.Ti017 = Convert.ToDecimal(dt.Rows[i]["Ti017"]);
							myObject.Ti018 = Convert.ToDecimal(dt.Rows[i]["Ti018"]);
							myObject.Ti019 = Convert.ToDecimal(dt.Rows[i]["Ti019"]);
							myObject.Ti020 = Convert.ToDecimal(dt.Rows[i]["Ti020"]);
							myObject.Ti021 = Convert.ToDecimal(dt.Rows[i]["Ti021"]);
							myObject.Ti022 = Convert.ToDecimal(dt.Rows[i]["Ti022"]);

							list.Add(myObject);
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("載入[ANBTI]資料時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return list;
			}
		}

		/// <summary>
		/// 取得部門(510)的加總彙總表
		/// </summary>
		/// <param name="year">年度</param>
		/// <returns></returns>
		public static DataTable LoadDept5_SumData(string year)
		{
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
					SQL.AppendFormat("SELECT TI003, TI004, TI006, TI007 AS '會科編號', TI008 AS '會科名稱', " +
						"TI009 AS '會科說明', SUM(TI010) AS TI010, SUM(TI011) AS TI011, SUM(TI012) AS TI012, SUM(TI013) AS TI013, " +
						"SUM(TI014) AS TI014, SUM(TI015) AS TI015, SUM(TI016) AS TI016, SUM(TI017) AS TI017, SUM(TI018) AS TI018, " +
						"SUM(TI019) AS TI019, SUM(TI020) AS TI020, SUM(TI021) AS TI021, SUM(TI022) AS TI022 " +
						"FROM dbo.ANBTI " +
						"WHERE TI002 like '5%' " +
						//"AND TI008 <> '總計' " +
						"AND TI003 = '{0}' " +
						"GROUP BY TI003, TI004, TI006, TI007, TI008, TI009 " +
						"ORDER BY TI006 ", year);

					if (conn.State != ConnectionState.Open)
						conn.Open();
					command = new SqlCommand(SQL.ToString(), conn);

					adapter = new SqlDataAdapter(command);

					adapter.Fill(dt);
					
				}
				catch (Exception e)
				{
					_log.Debug("載入[ANBTI]資料時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return dt;
				
			}
		}

		public static DataGridView SetDGV(DataGridView dgv, List<Object> list, DataTable tmplTable, bool isDeptAccounting, string Dept)
		{			
			DataTable table = tmplTable;

			try
			{
				if ((dgv.Rows.Count != table.Rows.Count) && isDeptAccounting == false)
					dgv = Set_dgvSummary(table, dgv, isDeptAccounting, Dept);	// 設定總表的RowHeaderCell


				if (list != null && list.Count > 0)
				{
					
					int listCount = list.Count;
					int rowCount = dgv.Rows.Count;

					if (rowCount < listCount)
						for (int i = 0; i < (listCount - rowCount) + 1; i++)
							dgv.Rows.Add();
					


					for (int i = 0; i < list.Count; i++)
					{

						ANBTI myObj = (ANBTI)list[i];
						string key = "";

						// 組HearerCell的標題文字(會科編號和名稱)
						if (String.IsNullOrEmpty(myObj.Ti007))
							key = myObj.Ti008 + "-(" + myObj.Ti009 + ")";
						else
							key = myObj.Ti007 + "-" + myObj.Ti008 + "-(" + myObj.Ti009 + ")";


						// 將數字填入表格
						for (int j = 0; j < dgv.RowCount; j++)
						{
						
							if (key.Equals(dgv.Rows[j].HeaderCell.Value.ToString()))
							{
								if (!"0".Equals(myObj.Ti010.ToString()))
									dgv.Rows[j].Cells[0].Value = myObj.Ti010;  // 1月

								if (!"0".Equals(myObj.Ti011.ToString()))
									dgv.Rows[j].Cells[1].Value = myObj.Ti011;  // 2月

								if (!"0".Equals(myObj.Ti012.ToString()))
									dgv.Rows[j].Cells[2].Value = myObj.Ti012;  // 3月

								if (!"0".Equals(myObj.Ti013.ToString()))
									dgv.Rows[j].Cells[3].Value = myObj.Ti013;  // 4月

								if (!"0".Equals(myObj.Ti014.ToString()))
									dgv.Rows[j].Cells[4].Value = myObj.Ti014;  // 5月

								if (!"0".Equals(myObj.Ti015.ToString()))
									dgv.Rows[j].Cells[5].Value = myObj.Ti015;  // 6月

								if (!"0".Equals(myObj.Ti016.ToString()))
									dgv.Rows[j].Cells[6].Value = myObj.Ti016;  // 7月

								if (!"0".Equals(myObj.Ti017.ToString()))
									dgv.Rows[j].Cells[7].Value = myObj.Ti017;  // 8月

								if (!"0".Equals(myObj.Ti018.ToString()))
									dgv.Rows[j].Cells[8].Value = myObj.Ti018;  // 9月

								if (!"0".Equals(myObj.Ti019.ToString()))
									dgv.Rows[j].Cells[9].Value = myObj.Ti019;  // 10月

								if (!"0".Equals(myObj.Ti020.ToString()))
									dgv.Rows[j].Cells[10].Value = myObj.Ti020; // 11月

								if (!"0".Equals(myObj.Ti021.ToString())) 
									dgv.Rows[j].Cells[11].Value = myObj.Ti021; // 12月

								if (!"0".Equals(myObj.Ti022.ToString()))
									dgv.Rows[j].Cells[12].Value = myObj.Ti022; // 項目總計
								
								dgv.Rows[j].Cells[13].Value = myObj.Ti000; // 唯一序號

								break;
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				_log.Debug("ANBTI_Model的SetDGV方式發生錯誤：" + ex.Message);
				MessageBox.Show("ANBTI_Model的SetDGV方式發生錯誤：" + ex.Message);
			}

			return dgv;
		}

		/// <summary>
		/// 設定總表的列的RowHeaderCell
		/// </summary>
		/// <param name="dt">包含樣版的DataTable</param>
		/// <param name="dgv_Summary">準備套用的目標DataGridView</param>
		/// <returns></returns>
		public static DataGridView Set_dgvSummary(DataTable dt, DataGridView dgv_Summary, bool isDeptAccounting, string Dept)
		{
			string[] DeptsForHideSpecRow = new string[]{ "510", "511", "513", "520" ,"710"};

			// 對應ANBTL的TL001，將含有「直接」、「人力」、「外勞」字眼的index搜出來，在總表中visible = false
			string[] SpecRow = new string[] { "直接", "人力", "外勞" };	


			_log.Debug("執行Set_dgvSummary");
			dgv_Summary.Rows.Clear();
			//int ManualCount = 0;	// 記錄財會幫忙輸入數值的列數

			if (dt != null) { 
			// 設定每一個會科項目
				for (int i = 0; i < dt.Rows.Count; i++)
				{
					string c1 = dt.Rows[i]["會科編號"].ToString();  // 會科編號
					string c2 = dt.Rows[i]["會科名稱"].ToString();  // 會科名稱
					string c3 = dt.Rows[i]["會科說明"].ToString();  // 會科名稱說明
				
					dgv_Summary.Rows.Add();

				
					if (String.IsNullOrEmpty(c1))
						dgv_Summary.Rows[i].HeaderCell.Value = c2 + "-(" + c3 + ")";
					else
						dgv_Summary.Rows[i].HeaderCell.Value = c1 + "-" + c2 + "-(" + c3 + ")";


					if (isDeptAccounting == true && !"M".Equals(dt.Rows[i]["TL010"].ToString()))	// 是否是財會部的人登入，且該欄位是否是由財會部協助填寫
						dgv_Summary.Rows[i].Visible = false;	


					// 若登入者是特定部門，則隱藏特定的Row
					if (DeptsForHideSpecRow.Contains(Dept) && SpecRow.Contains(c3))
						dgv_Summary.Rows[i].Visible = false;


					// 將需要預先填上的費用(水電、網路、瓦斯等)，加上Tag
					if (dt.Rows[i]["TL012"] != null && !String.IsNullOrEmpty(dt.Rows[i]["TL012"].ToString())) { 
						dgv_Summary.Rows[i].Tag = dt.Rows[i]["TL012"].ToString();
						dgv_Summary.Rows[i].ReadOnly = true;
					}

				}
			}

			dgv_Summary.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

			return dgv_Summary;
		}

		/// <summary>
		/// 計算預算總表的每月總計(最後一列)並回傳List
		/// </summary>
		/// <param name="dgv"></param>
		/// <param name="list"></param>
		/// <returns></returns>
		public  static List<Object> GetMonthlySum(DataGridView dgv, List<Object> list) {
			
			decimal sum = 0;
			try {

				if (dgv != null && dgv.RowCount > 0) 
				{
					for (int j = 0; j < 13; j++)	// 各個月份和項目總計
					{
						sum = 0;
						for (int i = 0; i < dgv.RowCount - 1; i++) // 最後一行(各月份總計)略過
						{	
							if(!dgv.Rows[i].HeaderCell.Value.ToString().Contains("伙食人數"))	// 伙食人數的值，不加到加總裡面
								sum += Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);   // 加總各項目個別月份的值
						}						

						dgv.Rows[dgv.RowCount - 1].Cells[j].Value = sum;
						list.Add(dgv.Rows[dgv.RowCount - 1].Cells[j].Value);
					}
				}
			}
			catch (Exception ex) {
				_log.Debug("計算每月總計時發生錯誤：" + ex.Message);
				MessageBox.Show("計算每月總計時發生錯誤：" + ex.Message);
			}

			return list;
		}

		/// <summary>
		/// 取得每一個部門(預算總表)的每月總計
		/// </summary>
		/// <param name="year"></param>
		/// <returns></returns>
		public static DataTable GetSumByDept(string year)
		{						
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
					SQL.AppendFormat("SELECT TI002, " +
						"SUM(TI010) AS TI010, " +
						"SUM(TI011) AS TI011, " +
						"SUM(TI012) AS TI012, " +
						"SUM(TI013) AS TI013, " +
						"SUM(TI014) AS TI014, " +
						"SUM(TI015) AS TI015, " +
						"SUM(TI016) AS TI016, " +
						"SUM(TI017) AS TI017, " +
						"SUM(TI018) AS TI018, " +
						"SUM(TI019) AS TI019, " +
						"SUM(TI020) AS TI020, " +
						"SUM(TI021) AS TI021, " +						
						"SUM(TI022) AS TI022 " +
						"FROM dbo.ANBTI " +
						"WHERE TI003 = '{0}' and TI007 <> '' " +
						"GROUP BY TI002 ", year);

					if (conn.State != ConnectionState.Open)
						conn.Open();
					command = new SqlCommand(SQL.ToString(), conn);

					adapter = new SqlDataAdapter(command);

					adapter.Fill(dt);
				}
				catch (Exception e)
				{
					_log.Debug("執行[GetSumByDept]方法時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return dt;
			}
		}

		/// <summary>
		/// 取得所有會計科目的每月總計
		/// </summary>
		/// <param name="year"></param>
		/// <returns></returns>
		public static DataTable GetSumByAccounts(string year)
		{
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
					SQL.AppendFormat("SELECT TI007, TI008, TI009, " +
						"SUM(TI010) AS TI010, " +
						"SUM(TI011) AS TI011, " +
						"SUM(TI012) AS TI012, " +
						"SUM(TI013) AS TI013, " +
						"SUM(TI014) AS TI014, " +
						"SUM(TI015) AS TI015, " +
						"SUM(TI016) AS TI016, " +
						"SUM(TI017) AS TI017, " +
						"SUM(TI018) AS TI018, " +
						"SUM(TI019) AS TI019, " +
						"SUM(TI020) AS TI020, " +
						"SUM(TI021) AS TI021, " +
						"SUM(TI022) AS TI022 " +
						"FROM dbo.ANBTI " +
						"WHERE TI003 = '{0}' and TI007 <> '' " +
						"GROUP BY TI007, TI008, TI009", year);

					if (conn.State != ConnectionState.Open)
						conn.Open();
					command = new SqlCommand(SQL.ToString(), conn);

					adapter = new SqlDataAdapter(command);

					adapter.Fill(dt);
				}
				catch (Exception e)
				{
					_log.Debug("執行[GetSumByAccounts]方法時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return dt;
			}
		}
	}
}
