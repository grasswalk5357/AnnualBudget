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
    class ANBTF_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// 新增或更新ANBTF的資料
		/// </summary>
		/// <param name="list">存放ANBTF物件的List</param>
		/// <returns></returns>
		public static bool UpdateToANBTF(List<Object> myList)
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
					foreach (CapEx myObj in myList) { 
						SQL_1.Clear();
						SQL_2.Clear();

						SQL_1.AppendFormat("IF EXISTS (SELECT TF000 FROM dbo.ANBTF WHERE TF000 = {0} ) " +
											"UPDATE dbo.ANBTF SET FLAG = FLAG + 1, MODIFIER = '{1}', MODI_DATE = '{2}', " +
											"TF004 = '{3}', TF005 = '{4}', TF006 = '{5}', TF007 = {6}, TF008 = {7}, " +
											"TF009 = {8}, TF010 = {9}, TF011 = {10}, TF012 = {11}, TF013 = {12}, " +
											"TF014 = {13}, TF015 = '{14}', TF016 = '{15}', TF017 = '{16}', TF018 = '{17}' " +
											"WHERE TF000 = {18}  ELSE ",
											myObj.Id, myObj.Modifier, myObj.Modi_date,
											myObj.AssetType, myObj.Name, myObj.Spec, myObj.Num, myObj.UnitPrice, 
											myObj.TotalPrice, myObj.Life, myObj.GetYear, myObj.GetMonth, myObj.GetDay, 
											myObj.Depre, myObj.Purpose, myObj.Benifit, myObj.Trade, myObj.IsDelete,
											myObj.Id);

						SQL_2.AppendFormat("INSERT INTO dbo.ANBTF " +
							"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, " +
							"TF001, TF002, TF003, TF004, TF005, TF006, " +
							"TF007, TF008, TF009, TF010, TF011, TF012, " +
							"TF013, TF014, TF015, TF016, TF017, TF018 ) " +
							"VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, " +
							"'{7}', '{8}', '{9}', '{10}', '{11}', " +
							"'{12}', {13}, {14}, {15}, {16}, " +
							"{17}, {18}, {19}, {20}, '{21}', " +
							"'{22}', '{23}', '{24}' " +
							" ) ",
							myObj.Company, myObj.Creator, myObj.Usr_group, myObj.Create_date, "", "", myObj.Flag,
							myObj.AnnualBudgetFormID, myObj.DeptNo, myObj.Year, myObj.AssetType, myObj.Name, 
							myObj.Spec, myObj.Num, myObj.UnitPrice, myObj.TotalPrice, myObj.Life, 
							myObj.GetYear, myObj.GetMonth, myObj.GetDay, myObj.Depre, myObj.Purpose, 
							myObj.Benifit, myObj.Trade, myObj.IsDelete);

						SQL_1 = SQL_1.Append(SQL_2);


						//command = new SqlCommand(SQL_1.ToString(), conn);
						command.CommandText = SQL_1.ToString();
						command.ExecuteNonQuery();
					}
					trans.Commit();

				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTF」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「ANBTF」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTF」時發生錯誤的訊息：" + e2.Message);
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
		/// 將資本支出預算表DataGridView的資料儲存至List
		/// </summary>
		/// <param name="dgv">資本支出預算表的DataGridView</param>
		/// <param name="AnnualBudgetFormID">資本支出預算表的代碼</param>
		/// <param name="MainDeptNo">編寫年度預算的目標部門代號</param>
		/// <param name="deptNo">登入者的部門代號</param>
		/// <param name="year">年份</param>
		/// <param name="userId">登入者的ID</param>
		/// <param name="modidate">修改日期</param>
		/// <returns></returns>
		public static List<Object> SaveMatrixData(DataGridView dgv, string AnnualBudgetFormID, string MainDeptNo, string deptNo, string year, string userId, string modidate)
        {
            List<Object> list = new List<Object>();
            CapEx myObject = null;
			try
			{
				for (int i = 0; i < dgv.RowCount - 1; i++)
				{
					myObject = new CapEx();
					myObject.Creator = userId;
					myObject.Usr_group = deptNo;
					myObject.Modifier = userId;
					myObject.Modi_date = modidate;
					myObject.DeptNo = MainDeptNo;
					myObject.Year = year;
					myObject.AnnualBudgetFormID = AnnualBudgetFormID;

					myObject.No = Convert.ToDecimal(dgv.Rows[i].HeaderCell.Value);

					if (dgv.Rows[i].Cells[0].Value != null)
						myObject.AssetType = dgv.Rows[i].Cells[0].Value.ToString();

					if (dgv.Rows[i].Cells[1].Value != null)
						myObject.Name = dgv.Rows[i].Cells[1].Value.ToString();  // 設備名稱

					if (dgv.Rows[i].Cells[2].Value != null)
						myObject.Spec = dgv.Rows[i].Cells[2].Value.ToString();  // 規格

					myObject.Num = Convert.ToDecimal(dgv.Rows[i].Cells[3].Value);           // 數量     
					myObject.UnitPrice = Convert.ToDecimal(dgv.Rows[i].Cells[4].Value);     // 單價


					myObject.TotalPrice = Convert.ToDecimal(dgv.Rows[i].Cells[5].Value);    // 總價金額
					myObject.Life = Convert.ToDecimal(dgv.Rows[i].Cells[6].Value);          // 耐用年限
					myObject.GetYear = Convert.ToDecimal(dgv.Rows[i].Cells[7].Value);       // 年
					myObject.GetMonth = Convert.ToDecimal(dgv.Rows[i].Cells[8].Value);      // 月
					myObject.GetDay = Convert.ToDecimal(dgv.Rows[i].Cells[9].Value);        // 日
					myObject.Depre = Convert.ToDecimal(dgv.Rows[i].Cells[10].Value);        // 每月折舊

					if (dgv.Rows[i].Cells[11].Value != null)
						myObject.Purpose = dgv.Rows[i].Cells[11].Value.ToString();          // 增設(改善)目的

					if (dgv.Rows[i].Cells[12].Value != null)
						myObject.Benifit = dgv.Rows[i].Cells[12].Value.ToString();          // 預期效益評估

					if (dgv.Rows[i].Cells[13].Value != null)
						myObject.Trade = dgv.Rows[i].Cells[13].Value.ToString();            // 交易對象

					DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dgv.Rows[i].Cells[14]; // 刪除標記
					if (chk.Value != null && (bool)chk.Value)
						myObject.IsDelete = "Y";
					else
						myObject.IsDelete = "N";

					myObject.Id = Convert.ToDecimal(dgv.Rows[i].Cells[15].Value);           // 唯一序號

					list.Add(myObject);
				}
			}		
			catch (Exception e)
			{
				_log.Debug("儲存[資本支出預算表]時發生錯誤：" + e.Message);
				MessageBox.Show("儲存[資本支出預算表]時發生錯誤：" + e.Message);
			}


			return list;
        }

		/// <summary>
		/// 計算每月攤提至List
		/// </summary>
		/// <param name="list">存放ANBTF物件的List</param>
		/// <param name="dgv">資本支出預算表的DataGridView</param>
		/// <returns></returns>
		public static List<decimal[]> GetMonthlyData(List<Object> list, DataGridView dgv)
        {
            List<decimal[]> dList = new List<decimal[]>();
            decimal[] monthlyData1 = new decimal[13];
            decimal[] monthlyData2 = new decimal[13];

            CapEx myObject = null;
			DataGridViewComboBoxCell cbc = null;

			List<string> tempList = new List<string>();

			if (dgv.Rows[0].Cells[0] != null) { 
				cbc = (DataGridViewComboBoxCell)dgv.Rows[0].Cells[0];
				for (int i = 0; i < cbc.Items.Count; i++) { 
					if (!tempList.Contains(cbc.Items[i].ToString()))
					{
						tempList.Add(cbc.Items[i].ToString());
					}
				}
			}



			decimal[,] monthlyData = new decimal[tempList.Count, 13];

			for (int j = 0; j < list.Count; j++)    // List的筆數
			{
				myObject = (CapEx)list[j];

				if ("N".Equals(myObject.IsDelete))  // 刪除標記沒有勾選的才計算
				{
					for (int i = 0; i < 12; i++)    // 月份
					{                
                    
						if (myObject.GetMonth <= (i + 1))
						{
							//monthlyData[tempList.IndexOf(myObject.Type), i] += myObject.Depre;
							//monthlyData[tempList.IndexOf(myObject.Type), 12] += myObject.Depre;

							if (tempList.IndexOf(myObject.AssetType) == 0)	// 固定資產
							{
								monthlyData1[i] += myObject.Depre;
								monthlyData1[12] += myObject.Depre;
							}
							else if (tempList.IndexOf(myObject.AssetType) == 1)     // 電腦軟體
							{
								monthlyData2[i] += myObject.Depre;
								monthlyData2[12] += myObject.Depre;
							}
						}
					}
                }
            }

            dList.Add(monthlyData1);
            dList.Add(monthlyData2);
            //return monthlyData;
            return dList;
        }

		/// <summary>
		/// 從DB載入ANBTF至List
		/// </summary>
		/// <param name="annualBudgetFormID">資本支出預算表的代碼</param>
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
					SQL.AppendFormat("SELECT TF000, TF001, TF002, TF003, TF004, " +
									 "TF005, TF006, TF007, TF008, TF009, " +
									 "TF010, TF011, TF012, TF013, TF014, " +
									 "TF015, TF016, TF017 " +
									 "FROM dbo.ANBTF " +
									 "WHERE TF001 = '{0}' AND TF002 = '{1}' " +
									 "AND TF003 = '{2}' AND TF018 = 'N' ", annualBudgetFormID, deptNo, year);

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
							ANBTF myObject = new ANBTF();
							myObject.Tf000 = Convert.ToDecimal(dt.Rows[i]["TF000"]);   // 流水號
							myObject.Tf001 = dt.Rows[i]["TF001"].ToString();           // 單別代號
							myObject.Tf002 = dt.Rows[i]["TF002"].ToString();           // 部門代號
							myObject.Tf003 = dt.Rows[i]["TF003"].ToString();           // 年度
							myObject.Tf004 = dt.Rows[i]["TF004"].ToString();
							myObject.Tf005 = dt.Rows[i]["TF005"].ToString();
							myObject.Tf006 = dt.Rows[i]["TF006"].ToString();
							myObject.Tf007 = Convert.ToDecimal(dt.Rows[i]["TF007"]);
							myObject.Tf008 = Convert.ToDecimal(dt.Rows[i]["TF008"]);
							myObject.Tf009 = Convert.ToDecimal(dt.Rows[i]["TF009"]);
							myObject.Tf010 = Convert.ToDecimal(dt.Rows[i]["TF010"]);
							myObject.Tf011 = Convert.ToDecimal(dt.Rows[i]["TF011"]);
							myObject.Tf012 = Convert.ToDecimal(dt.Rows[i]["TF012"]);
							myObject.Tf013 = Convert.ToDecimal(dt.Rows[i]["TF013"]);
							myObject.Tf014 = Convert.ToDecimal(dt.Rows[i]["TF014"]);
							myObject.Tf015 = dt.Rows[i]["TF015"].ToString();
							myObject.Tf016 = dt.Rows[i]["TF016"].ToString();
							myObject.Tf017 = dt.Rows[i]["TF017"].ToString();


							list.Add(myObject);
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("載入[ANBTF]資料時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return list;
			}
		}

		/// <summary>
		/// 將從ANBTF載入的資料帶入DataGridView裡面
		/// </summary>
		/// <param name="dgv">欲接受資料的DataGridView</param>
		/// <param name="list">存放ANBTF物件的List</param>
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
						ANBTF myObj = (ANBTF)list[i];

						cbc = (DataGridViewComboBoxCell)dgv.Rows[i].Cells[0];
						for (int j = 0; j < cbc.Items.Count; j++)
							if (cbc.Items[j].ToString().Equals(myObj.Tf004.ToString()))
								cbc.Value = cbc.Items[j];           // 資產類別

						//dgv.Rows[i].Cells[0].Value = myObj.Tf004;   
						dgv.Rows[i].Cells[1].Value = myObj.Tf005;   // 設備名稱
						dgv.Rows[i].Cells[2].Value = myObj.Tf006;   // 規格
						dgv.Rows[i].Cells[3].Value = myObj.Tf007;   // 數量
						dgv.Rows[i].Cells[4].Value = myObj.Tf008;   // 單價
						dgv.Rows[i].Cells[5].Value = myObj.Tf009;   // 總價金額
						dgv.Rows[i].Cells[6].Value = myObj.Tf010;   // 耐用年限

						dgv.Rows[i].Cells[7].Value = myObj.Tf011;   // 預定取得(年)

						//dgv.Rows[i].Cells[8].Value = myObj.Tf012; 
						cbc = (DataGridViewComboBoxCell)dgv.Rows[i].Cells[8];
						for (int j = 0; j < cbc.Items.Count; j++)
							if (cbc.Items[j].ToString().Equals(myObj.Tf012.ToString()))
								cbc.Value = cbc.Items[j];           // 預定取得(月)



						dgv.Rows[i].Cells[9].Value = myObj.Tf013;   // 預定取得(日)
						dgv.Rows[i].Cells[10].Value = myObj.Tf014;  // 每月折舊			
						dgv.Rows[i].Cells[11].Value = myObj.Tf015;  // 增設(改善)目的
						dgv.Rows[i].Cells[12].Value = myObj.Tf016;  // 預期效益評估
						dgv.Rows[i].Cells[13].Value = myObj.Tf017;  // 交易對象

						dgv.Rows[i].Cells[15].Value = myObj.Tf000;  // 唯一序號

					}
				}
			}
			catch (Exception ex)
			{
				_log.Debug("ANBTF_Model的SetDGV方式發生錯誤：" + ex.Message);
				MessageBox.Show("ANBTF_Model的SetDGV方式發生錯誤：" + ex.Message);
			}

			return dgv;
		}
	}
}
