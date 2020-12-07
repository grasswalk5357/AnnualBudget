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
    class ANBTE_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// 新增或更新資料至ANBTE資料表
		/// </summary>
		/// <param name="myList">存放ANBTE物件的List</param>
		/// <returns></returns>
		public static bool UpdateToANBTE(List<Object> myList)
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
					foreach (BizTrip myObj in myList) { 
						SQL_1.Clear();
						SQL_2.Clear();

						SQL_1.AppendFormat("IF EXISTS (SELECT TE000 FROM dbo.ANBTE WHERE TE000 = {0}) " +
												"UPDATE dbo.ANBTE SET FLAG = FLAG + 1, MODIFIER = '{1}', MODI_DATE = '{2}', " +
												"TE004 = {3}, TE005 = '{4}', TE006 = '{5}', TE007 = {6}, TE008 = {7}, " +
												"TE009 = {8}, TE010 = {9}, TE011 = {10}, TE012 = {11}, TE013 = {12}, " +
												"TE014 = {13}, TE015 = {14}, TE016 = {15}, TE017 = '{16}', TE018 = '{17}' " +
												"WHERE TE000 = {18}  ELSE ",
												myObj.Id, myObj.Modifier, myObj.Modi_date,
												myObj.Month, myObj.Name, myObj.Location, myObj.Days, myObj.AirFare, 
												myObj.HotelFare, myObj.ShippingExpenses, myObj.OtherFare, myObj.DailyExpense, myObj.FoodStipend, 
												myObj.SumOfTripExpense, myObj.Entertainment, myObj.TripInsurance, myObj.Memo, myObj.IsDelete, 
												myObj.Id);

						SQL_2.AppendFormat("INSERT INTO dbo.ANBTE " +
							"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, " +
							"TE001, TE002, TE003, TE004, TE005, TE006, " +
							"TE007, TE008, TE009, TE010, TE011, TE012, " +
							"TE013, TE014, TE015, TE016, TE017, TE018 ) " +
							"VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, " +
							"'{7}', '{8}', '{9}', {10}, '{11}', " +
							"'{12}', {13}, {14}, {15}, {16}, " +
							"{17}, {18}, {19}, {20}, {21}, " +
							"{22}, '{23}', '{24}' " +
							" ) ",
							myObj.Company, myObj.Creator, myObj.Usr_group, myObj.Create_date, "", "", myObj.Flag,
							myObj.AnnualBudgetFormID, myObj.DeptNo, myObj.Year, myObj.Month, myObj.Name, 
							myObj.Location, myObj.Days, myObj.AirFare, myObj.HotelFare, myObj.ShippingExpenses, 
							myObj.OtherFare, myObj.DailyExpense, myObj.FoodStipend, myObj.SumOfTripExpense, myObj.Entertainment, 
							myObj.TripInsurance, myObj.Memo, myObj.IsDelete);

						SQL_1 = SQL_1.Append(SQL_2);
																		
						command.CommandText = SQL_1.ToString();
						command.ExecuteNonQuery();
					}
					trans.Commit();

				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTE」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「ANBTE」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTE」時發生錯誤的訊息：" + e2.Message);
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
		/// 將出差計劃表DataGridView的資料儲存至List
		/// </summary>
		/// <param name="dgv">出差計劃表的DataGridView</param>
		/// <param name="AnnualBudgetFormID">出差計劃表的代碼</param>
		/// <param name="mainDeptNo">編寫年度預算的目標部門代號</param>
		/// <param name="deptNo">登入者的部門代號</param>
		/// <param name="year">年份</param>
		/// <param name="userId">登入者的ID</param>
		/// <param name="modidate">修改日期</param>
		/// <returns></returns>
		public static List<Object> SaveMatrixData(DataGridView dgv, string AnnualBudgetFormID,string MainDeptNo, string deptNo, string year, string userId, string modidate)
        {
            List<Object> list = new List<object>();

            BizTrip myObject;
			try
			{
				for (int i = 0; i < dgv.RowCount - 1; i++)
				{
					myObject = new BizTrip();
					myObject.Creator = userId;
					myObject.Usr_group = deptNo;
					myObject.Modifier = userId;
					myObject.Modi_date = modidate;
					myObject.DeptNo = MainDeptNo;
					myObject.Year = year;
					myObject.AnnualBudgetFormID = AnnualBudgetFormID;

					myObject.No = Convert.ToDecimal(dgv.Rows[i].HeaderCell.Value);
					myObject.Month = Convert.ToDecimal(dgv.Rows[i].Cells[0].Value);				// 月份

					if (dgv.Rows[i].Cells[1].Value != null)
						myObject.Name = dgv.Rows[i].Cells[1].Value.ToString();					// 出差者

					if (dgv.Rows[i].Cells[2].Value != null)
						myObject.Location = dgv.Rows[i].Cells[2].Value.ToString();				// 出差地區

					myObject.Days = Convert.ToDecimal(dgv.Rows[i].Cells[3].Value);				// 天數
					myObject.AirFare = Convert.ToDecimal(dgv.Rows[i].Cells[4].Value);			// 機票款
					myObject.HotelFare = Convert.ToDecimal(dgv.Rows[i].Cells[5].Value);			// 住宿費
					myObject.ShippingExpenses = Convert.ToDecimal(dgv.Rows[i].Cells[6].Value);	// 交通費
					myObject.OtherFare = Convert.ToDecimal(dgv.Rows[i].Cells[7].Value);			// 雜費
					myObject.DailyExpense = Convert.ToDecimal(dgv.Rows[i].Cells[8].Value);		// 日支費
					myObject.FoodStipend = Convert.ToDecimal(dgv.Rows[i].Cells[9].Value);		// 餐費
					myObject.SumOfTripExpense = Convert.ToDecimal(dgv.Rows[i].Cells[10].Value);	// 旅費小計
					myObject.Entertainment = Convert.ToDecimal(dgv.Rows[i].Cells[11].Value);	// 交際費
					myObject.TripInsurance = Convert.ToDecimal(dgv.Rows[i].Cells[12].Value);	// 旅平險
					if (dgv.Rows[i].Cells[13].Value != null)	
						myObject.Memo = dgv.Rows[i].Cells[13].Value.ToString();					// 出差目的說明

					// 標記刪除
					DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dgv.Rows[i].Cells[14];
					if (chk.Value != null && (bool)chk.Value)
						myObject.IsDelete = "Y";
					else
						myObject.IsDelete = "N";

					myObject.Id = Convert.ToDecimal(dgv.Rows[i].Cells[15].Value);				// 唯一序號

					list.Add(myObject);
				}
			}
			catch (Exception e) {
				_log.Debug("儲存[出差計劃表]時發生錯誤：" + e.Message);
				MessageBox.Show("儲存[出差計劃表]時發生錯誤：" + e.Message);
			}

			return list;
        }

		/// <summary>
		/// 計算每月攤提至List
		/// </summary>
		/// <param name="list">存放ANBTE物件的List</param>
		/// <returns></returns>
		public static List<decimal[]> GetMonthlyData(List<Object> list)
        {
            List<decimal[]> result = new List<decimal[]>();
            decimal[] monthlyData_TripExpense = new decimal[13];
            decimal[] monthlyData_Entertainment = new decimal[13];
            decimal[] monthlyData_Insurance = new decimal[13];

            BizTrip myObject = null;

			for (int j = 0; j < list.Count; j++)	// 資料筆數
            {
                myObject = (BizTrip)list[j];
				if ("N".Equals(myObject.IsDelete))  // 刪除標記沒有勾選的才計算
				{
					for (int i = 0; i < 12; i++)    // 月份
					{
						if (myObject.Month == (i + 1))
						{
							monthlyData_TripExpense[i] += myObject.SumOfTripExpense;
							monthlyData_TripExpense[12] += myObject.SumOfTripExpense;

							monthlyData_Entertainment[i] += myObject.Entertainment;
							monthlyData_Entertainment[12] += myObject.Entertainment;

							monthlyData_Insurance[i] += myObject.TripInsurance;
							monthlyData_Insurance[12] += myObject.TripInsurance;
						}
					}
                }
            }
            result.Add(monthlyData_TripExpense);    // 差旅費
            result.Add(monthlyData_Entertainment);  // 交際費
            result.Add(monthlyData_Insurance);      // 旅平險

            return result;
        }

		/// <summary>
		/// 從DB載入ANBTE至List
		/// </summary>
		/// <param name="annualBudgetFormID">出差計劃表的代碼</param>
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
					SQL.AppendFormat("SELECT TE000, TE001, TE002, TE003, TE004, " +
									 "TE005, TE006, TE007, TE008, TE009, " +
									 "TE010, TE011, TE012, TE013, TE014, " +
									 "TE015, TE016, TE017 " +
									 "FROM dbo.ANBTE " +
									 "WHERE TE001 = '{0}' AND TE002 = '{1}' " +
									 "AND TE003 = '{2}' AND TE018 = 'N' ", annualBudgetFormID, deptNo, year);

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
							ANBTE myObject = new ANBTE();
							myObject.Te000 = Convert.ToDecimal(dt.Rows[i]["TE000"]);   // 流水號
							myObject.Te001 = dt.Rows[i]["TE001"].ToString();           // 單別代號
							myObject.Te002 = dt.Rows[i]["TE002"].ToString();           // 部門代號
							myObject.Te003 = dt.Rows[i]["TE003"].ToString();           // 年度
							myObject.Te004 = Convert.ToDecimal(dt.Rows[i]["TE004"]);
							myObject.Te005 = dt.Rows[i]["TE005"].ToString();
							myObject.Te006 = dt.Rows[i]["TE006"].ToString();
							myObject.Te007 = Convert.ToDecimal(dt.Rows[i]["TE007"]);
							myObject.Te008 = Convert.ToDecimal(dt.Rows[i]["TE008"]);
							myObject.Te009 = Convert.ToDecimal(dt.Rows[i]["TE009"]);
							myObject.Te010 = Convert.ToDecimal(dt.Rows[i]["TE010"]);
							myObject.Te011 = Convert.ToDecimal(dt.Rows[i]["TE011"]);
							myObject.Te012 = Convert.ToDecimal(dt.Rows[i]["TE012"]);
							myObject.Te013 = Convert.ToDecimal(dt.Rows[i]["TE013"]);
							myObject.Te014 = Convert.ToDecimal(dt.Rows[i]["TE014"]);
							myObject.Te015 = Convert.ToDecimal(dt.Rows[i]["TE015"]);
							myObject.Te016 = Convert.ToDecimal(dt.Rows[i]["TE016"]);
							myObject.Te017 = dt.Rows[i]["TE017"].ToString();


							list.Add(myObject);
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("載入[ANBTE]資料時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return list;
			}
		}

		/// <summary>
		/// 將從ANBTE載入的資料帶入DataGridView裡面
		/// </summary>
		/// <param name="dgv">欲接受資料的DataGridView</param>
		/// <param name="list">存放ANBTE物件的List</param>
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

					if (rowCount <= listCount)
						for (int i = 0; i < (listCount - rowCount) + 1; i++)
							dgv.Rows.Add();

					for (int i = 0; i < list.Count; i++)
					{
						ANBTE myObj = (ANBTE)list[i];

						cbc = (DataGridViewComboBoxCell)dgv.Rows[i].Cells[0];
						for (int j = 0; j < cbc.Items.Count; j++)
							if (cbc.Items[j].ToString().Equals(myObj.Te004.ToString()))
								cbc.Value = cbc.Items[j];           // 月份

						//dgv.Rows[i].Cells[0].Value = myObj.Te004;   
						dgv.Rows[i].Cells[1].Value = myObj.Te005;   // 出差者
						dgv.Rows[i].Cells[2].Value = myObj.Te006;   // 出差地區
						dgv.Rows[i].Cells[3].Value = myObj.Te007;   // 天數
						dgv.Rows[i].Cells[4].Value = myObj.Te008;   // 機票款
						dgv.Rows[i].Cells[5].Value = myObj.Te009;   // 住宿費
						dgv.Rows[i].Cells[6].Value = myObj.Te010;   // 交通費

						dgv.Rows[i].Cells[7].Value = myObj.Te011;   // 雜費
						dgv.Rows[i].Cells[8].Value = myObj.Te012;   // 日支費
						dgv.Rows[i].Cells[9].Value = myObj.Te013;   // 餐費
						dgv.Rows[i].Cells[10].Value = myObj.Te014;  // 旅費小計			
						dgv.Rows[i].Cells[11].Value = myObj.Te015;  // 交際費
						dgv.Rows[i].Cells[12].Value = myObj.Te016;  // 旅平險
						dgv.Rows[i].Cells[13].Value = myObj.Te017;  // 出差目的說明

						dgv.Rows[i].Cells[15].Value = myObj.Te000;  // 唯一序號

					}
				}
			}
			catch (Exception ex) {
				_log.Debug("ANBTE_Model的SetDGV方式發生錯誤：" + ex.Message);
				MessageBox.Show("ANBTE_Model的SetDGV方式發生錯誤：" + ex.Message);
			}

			return dgv;
		}
	}
}
