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
    class ANBTG_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		//public static bool UpdateToANBTB(string deptNo, string Year, string AnnualBudgetFormID, string AnnualBudgetFormName, string isFormFilled)
		/// <summary>		
		/// 新增或更新ANBTG的資料
		/// </summary>
		/// <param name="list">存放ANBTG物件的List</param>
		/// <returns></returns>		
		public static bool UpdateToANBTG(Dictionary<string, Object> list)
		{
			
			bool result1 = true;
			
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
					
					for (int i = 0; i < list.Count; i++)
					{	
						ANBTG myObj = (ANBTG)list.Values.ElementAt(i);
						SQL_1.Clear();
						SQL_2.Clear();

						SQL_1.AppendFormat("IF EXISTS (SELECT TG000 FROM dbo.ANBTG WHERE TG000 = {0} ) " +
											"UPDATE dbo.ANBTG SET FLAG = FLAG + 1, MODIFIER = '{1}', MODI_DATE = '{2}', " +
											"TG004 = '{3}', TG005 = '{4}', TG006 = '{5}', TG007 = '{6}', TG008 = '{7}', " +
											"TG009 = '{8}', TG010 = '{9}', TG011 = '{10}', TG012 = {11}, TG013 = '{12}' " +											
											"WHERE TG000 = {13}  ELSE ",
											myObj.Tg000, myObj.Modifier, myObj.Modi_date,
											myObj.Tg004, myObj.Tg005, myObj.Tg006, myObj.Tg007, myObj.Tg008, 
											myObj.Tg009, myObj.Tg010, myObj.Tg011, myObj.Tg012, myObj.Tg013,
											myObj.Tg000);

						SQL_2.AppendFormat("INSERT INTO dbo.ANBTG " +
							"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, " +
							"TG001, TG002, TG003, TG004, TG005, TG006, " +
							"TG007, TG008, TG009, TG010, TG011, TG012, " +
							"TG013 ) " +
							"VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, " +
							"'{7}', '{8}', '{9}', '{10}', '{11}', " +
							"'{12}', '{13}', '{14}', '{15}', '{16}', " +
							"'{17}', {18}, '{19}' " +							
							" ) ",
							myObj.Company, myObj.Creator, myObj.Usr_group, myObj.Create_date, "", "", myObj.Flag,
							myObj.Tg001, myObj.Tg002, myObj.Tg003, myObj.Tg004, myObj.Tg005, 
							myObj.Tg006, myObj.Tg007, myObj.Tg008, myObj.Tg009, myObj.Tg010, 
							myObj.Tg011, myObj.Tg012, myObj.Tg013);

						SQL_1 = SQL_1.Append(SQL_2);

												
						command.CommandText = SQL_1.ToString();
						command.ExecuteNonQuery();


						if (ANBTH_Model.UpdateToANBTH(myObj.ANBTH_List) == false)
							throw new Exception();
					}
					trans.Commit();

				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTG」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「ANBTG」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTG」時發生錯誤的訊息：" + e2.Message);
					}
					result1 = false;
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return result1;
			}
		}


		/// <summary>
		/// 將專案彙總表單頭及單身的DataGridView資料儲存至Dictionary
		/// </summary>
		/// <param name="keyValues"></param>
		/// <param name="dgv_content">單身的DataGridView</param>
		/// <param name="row">目前點擊到的單頭的DataGridViewRow</param>
		/// <param name="AnnualBudgetFormID"></param>
		/// <param name="MainDeptNo">目前正在編寫年度預算的部門代號</param>
		/// <param name="deptNo">登入者的部門代號</param>
		/// <param name="year">年度</param>
		/// <param name="userId">登入者工號</param>
		/// <param name="modidate">修改日期</param>
		/// <returns></returns>
		public static Dictionary<string, Object> SaveMatrixData(Dictionary<string, Object> keyValues, DataGridView dgv_content, DataGridViewRow row, string AnnualBudgetFormID, string MainDeptNo, string deptNo, string year, string userId, string modidate)
        {
            string rdpid = "";
			ANBTG myObject = null;	// ANBTG儲存單頭
			ANBTH anbth = null;		// ANBTH儲存單身

			try
			{
				if (row.Cells[0].Value != null)
					rdpid = row.Cells[0].Value.ToString(); //記錄所選取的研發專案序號


				myObject = new ANBTG(rdpid);   // 將研發專案序號帶給ANBTG物件

				myObject.Creator = userId;
				myObject.Usr_group = deptNo;

				myObject.Modifier = userId;
				myObject.Modi_date = modidate;

				myObject.Tg001 = AnnualBudgetFormID;    // 表格代號  
				myObject.Tg002 = MainDeptNo;            // 部門代號
				myObject.Tg003 = year;                  // 年度			

				//rdp.RDP_Id = row.Cells[0].Value.ToString();         // 專案序號
				if (row.Cells[1].Value != null)
					myObject.Tg005 = row.Cells[1].Value.ToString();    // 機型

				if (row.Cells[2].Value != null)
					myObject.Tg006 = row.Cells[2].Value.ToString();    // 客戶
				if (row.Cells[3].Value != null)
					myObject.Tg007 = row.Cells[3].Value.ToString();    // 預計開案
				if (row.Cells[4].Value != null)
					myObject.Tg008 = row.Cells[4].Value.ToString();    // T/R預排日
				if (row.Cells[5].Value != null)
					myObject.Tg009 = row.Cells[5].Value.ToString();    // 預計結案
				if (row.Cells[6].Value != null)
					myObject.Tg010 = row.Cells[6].Value.ToString();    // 量產地
				if (row.Cells[7].Value != null)
					myObject.Tg011 = row.Cells[7].Value.ToString();    // 類型

				myObject.Tg012 = Convert.ToDecimal(row.Cells[8].Value);       // 耐用年限

				DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[9];	// 標記刪除

				if (chk.Value != null && (bool)chk.Value)
					myObject.Tg013 = "Y";
				else
					myObject.Tg013 = "N";

				myObject.Tg000 = Convert.ToDecimal(row.Cells[10].Value);

			}
			catch (Exception e) {
				_log.Debug("儲存[研發專案彙總表-單頭]時發生錯誤：" + e.Message);
				MessageBox.Show("儲存[研發專案彙總表-單頭]時發生錯誤：" + e.Message);
			}

			// 儲存單身 dgv_content裡面的值-------------------------------------------------------------------------
			try
			{
				for (int i = 0; i < dgv_content.Rows.Count; i++)
				{
					//ANBTH anbth = (ANBTH)myObject.RDP_Content_List[i];
					anbth = (ANBTH)myObject.ANBTH_List[i];

					anbth.Modifier = userId;
					anbth.Modi_date = modidate;

					anbth.Th001 = AnnualBudgetFormID;
					anbth.Th002 = MainDeptNo;
					anbth.Th003 = year;
					//myObject.RDP_Content_List[i].RDP_Id = "";
					anbth.Th005 = (i + 1).ToString();
					anbth.Th006 = dgv_content.Rows[i].HeaderCell.Value.ToString();


					for (int j = 0; j < 12; j++) { 
						if (dgv_content.Rows[i].Cells[j].Value != null && !String.IsNullOrEmpty(dgv_content.Rows[i].Cells[j].Value.ToString())) { 
							anbth.MonthlyData[j] = Convert.ToDecimal(dgv_content.Rows[i].Cells[j].Value);

							if ("8".Equals(anbth.Th005)) {
								
								// 預設耐用年限4年*12月 = 48，除以48
								anbth.Th024 = Math.Round(anbth.MonthlyData[j] / 48, 0, MidpointRounding.AwayFromZero);
							}
						}
						else
							anbth.MonthlyData[j] = 0;
					}

					anbth.Th019 = Convert.ToDecimal(dgv_content.Rows[i].Cells[12].Value);  // 總金額
					anbth.Th020 = Convert.ToDecimal(dgv_content.Rows[i].Cells[13].Value);  // 明躍支出
					anbth.Th021 = Convert.ToDecimal(dgv_content.Rows[i].Cells[14].Value);  // 立和支出
					anbth.Th022 = Convert.ToDecimal(dgv_content.Rows[i].Cells[15].Value);  // 客戶支出

					anbth.Th000 = Convert.ToDecimal(dgv_content.Rows[i].Cells[16].Value);  // 唯一序號

					if ("Y".Equals(myObject.Tg013))
						anbth.Th023 = "Y";
					else
						anbth.Th023 = "";
				}



				// 以單頭序號作為key值，儲存單頭(含單身)的物件進MAP
				if (keyValues.ContainsKey(rdpid))
					keyValues[rdpid] = myObject;
				else
					keyValues.Add(rdpid, myObject);
			}
			catch (Exception e) {
				_log.Debug("儲存[研發專案彙總表-單身]時發生錯誤：" + e.Message);
				MessageBox.Show("儲存[研發專案彙總表-單身]時發生錯誤：" + e.Message);
			}

            return keyValues;
        }

		/// <summary>
		/// 計算每月攤提至List[]
		/// </summary>
		/// <param name="myMap">存放ANBTG物件的Map</param>		
		/// <returns></returns>
		
		public static List<decimal[]> GetMonthlyData(Dictionary<string, Object> myMap)
        {			
			List<decimal[]> dList = new List<decimal[]>();
			
            decimal[] monthlyData1 = new decimal[13];
            decimal[] monthlyData2 = new decimal[13];
            decimal[] monthlyData3 = new decimal[13];
            decimal[] monthlyData4 = new decimal[13];
            decimal[] monthlyData5 = new decimal[13];
            decimal[] monthlyData6 = new decimal[13];
            decimal[] monthlyData7 = new decimal[13];
            decimal[] monthlyData8 = new decimal[13];
            decimal[] monthlyData9 = new decimal[13];
			decimal[] monthlyData10 = new decimal[13];
			decimal[] tempMonthlyData = new decimal[13];
			decimal tempNum = 0;

			ANBTG myObject = null;

			List<Object> list = myMap.Values.ToList();	// 專案的個數

            
            for (int i = 0; i < list.Count; i++)
            {
				myObject = (ANBTG)list[i];

				if ("N".Equals(myObject.Tg013))  // 刪除標記沒有勾選的才計算
				{
					ANBTH anbth = null;
					for (int y = 0; y < 12; y++)    // 月份 1~12
					{
						anbth = (ANBTH)myObject.ANBTH_List[0];
						monthlyData1[y] += anbth.MonthlyData[y];	//(1)加班費
						monthlyData1[12] += anbth.MonthlyData[y];

						anbth = (ANBTH)myObject.ANBTH_List[1];
						monthlyData2[y] += anbth.MonthlyData[y];	//(2)樣品費(消耗器材及原料使用費)
						monthlyData2[12] += anbth.MonthlyData[y];

						anbth = (ANBTH)myObject.ANBTH_List[2];
						monthlyData2[y] += anbth.MonthlyData[y];    //(3)樣品費(Mockup) → 併入(2)樣品費(消耗器材及原料使用費)一起計算
						monthlyData2[12] += anbth.MonthlyData[y];

						anbth = (ANBTH)myObject.ANBTH_List[3];
						monthlyData4[y] += anbth.MonthlyData[y];	//(4)技術移轉、設計費(勞務費)
						monthlyData4[12] += anbth.MonthlyData[y];

						anbth = (ANBTH)myObject.ANBTH_List[4];
						monthlyData5[y] += anbth.MonthlyData[y];	//(5)運費
						monthlyData5[12] += anbth.MonthlyData[y];

						anbth = (ANBTH)myObject.ANBTH_List[5];
						monthlyData6[y] += anbth.MonthlyData[y];	//(6)認證費
						monthlyData6[12] += anbth.MonthlyData[y];

						anbth = (ANBTH)myObject.ANBTH_List[6];
						monthlyData7[y] += anbth.MonthlyData[y];    //(7)模具費(未滿8萬元)
						monthlyData7[12] += anbth.MonthlyData[y];

						anbth = (ANBTH)myObject.ANBTH_List[7];
						monthlyData8[y] += anbth.MonthlyData[y];	//(8)模具(金額>=8萬元且耐用2年以上)
						monthlyData8[12] += anbth.MonthlyData[y];

						tempMonthlyData[y] += Math.Round(anbth.MonthlyData[y] / 48, 0, MidpointRounding.AwayFromZero);
						tempMonthlyData[12] += Math.Round(anbth.MonthlyData[y] / 48, 0, MidpointRounding.AwayFromZero);

						/*
						anbth = (ANBTH)myObject.ANBTH_List[8];
						monthlyData9[y] += anbth.MonthlyData[y];
						monthlyData9[12] += anbth.MonthlyData[y];
						*/
						anbth = (ANBTH)myObject.ANBTH_List[8];
						monthlyData7[y] += anbth.MonthlyData[y];	//(9)治具費(未滿8萬元) → 併入(7)模具費(未滿8萬元)一起計算
						monthlyData7[12] += anbth.MonthlyData[y];
					}					
				}
            }

			for (int i = 0; i < 12; i++) {
				monthlyData10[i] = tempNum + tempMonthlyData[i];
				tempNum = monthlyData10[i];
				monthlyData10[12] += monthlyData10[i];
				
			}

            dList.Add(monthlyData1);
            dList.Add(monthlyData2);
            dList.Add(monthlyData3);
            dList.Add(monthlyData4);
            dList.Add(monthlyData5);
            dList.Add(monthlyData6);
            dList.Add(monthlyData7);
            dList.Add(monthlyData8);
            dList.Add(monthlyData9);
			dList.Add(monthlyData10);

			return dList;
        }


		public static Dictionary<string, Object> LoadData(string annualBudgetFormID, string deptNo, string year)
		{
			//List<Object> list = null;
			Dictionary<string, Object> keyValues = null;
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
					SQL.AppendFormat("SELECT TG000, TG001, TG002, TG003, TG004, " +
									 "TG005, TG006, TG007, TG008, TG009, " +
									 "TG010, TG011, TG012 " +
									 "FROM dbo.ANBTG " +
									 "WHERE TG001 = '{0}' AND TG002 = '{1}' " +
									 "AND TG003 = '{2}' AND TG013 = 'N' ", annualBudgetFormID, deptNo, year);

					if (conn.State != ConnectionState.Open)
						conn.Open();
					command = new SqlCommand(SQL.ToString(), conn);

					adapter = new SqlDataAdapter(command);

					adapter.Fill(dt);

					if (dt.Rows.Count > 0)
					{

						//list = new List<object>();
						keyValues = new Dictionary<string, object>();
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							ANBTG myObject = new ANBTG(dt.Rows[i]["TG004"].ToString());
							myObject.Tg000 = Convert.ToDecimal(dt.Rows[i]["TG000"]);   // 流水號
							myObject.Tg001 = dt.Rows[i]["TG001"].ToString();           // 單別代號
							myObject.Tg002 = dt.Rows[i]["TG002"].ToString();           // 部門代號
							myObject.Tg003 = dt.Rows[i]["TG003"].ToString();           // 年度
							//myObject.Tg004 = dt.Rows[i]["TG004"].ToString();
							myObject.Tg005 = dt.Rows[i]["TG005"].ToString();
							myObject.Tg006 = dt.Rows[i]["TG006"].ToString();
							myObject.Tg007 = dt.Rows[i]["TG007"].ToString();
							myObject.Tg008 = dt.Rows[i]["TG008"].ToString();
							myObject.Tg009 = dt.Rows[i]["TG009"].ToString();
							myObject.Tg010 = dt.Rows[i]["TG010"].ToString();
							myObject.Tg011 = dt.Rows[i]["TG011"].ToString();
							myObject.Tg012 = Convert.ToDecimal(dt.Rows[i]["TG012"]);

							//list.Add(myObject);

							if (keyValues.ContainsKey(myObject.Tg004))
								keyValues[myObject.Tg004] = myObject;
							else
								keyValues.Add(myObject.Tg004, myObject);
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("載入[ANBTG]資料時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return keyValues;
			}
		}

		/// <summary>
		/// 將從ANBTG載入的資料帶入DataGridView裡面
		/// </summary>
		/// <param name="dgv">欲接受資料的DataGridView</param>
		/// <param name="list">存放ANBTG物件的List</param>
		/// <returns></returns>
		public static DataGridView SetDGV(DataGridView dgv, Dictionary<string, Object> list)
		{			
			try
			{
				dgv.Rows.Clear();
				if (list != null && list.Count > 0)
				{
					int listCount = list.Count;
					int rowCount = dgv.Rows.Count;

					if (rowCount < listCount)
					{
						for (int i = 0; i < (listCount - rowCount); i++) {
							dgv.Rows.Add();
						}

						dgv.RowHeadersWidth = 63;
					}


					for (int i = 0; i < list.Count; i++)
					{
						Object obj = list.Values.ElementAt(i);
						ANBTG myObj = (ANBTG)obj;

						dgv.Rows[i].Cells[0].Value = myObj.Tg004;   // 專案序號
						dgv.Rows[i].Cells[1].Value = myObj.Tg005;   // 機型
						dgv.Rows[i].Cells[2].Value = myObj.Tg006;   // 客戶
						dgv.Rows[i].Cells[3].Value = myObj.Tg007;   // 預計開案
						dgv.Rows[i].Cells[4].Value = myObj.Tg008;   // TR預排日
						dgv.Rows[i].Cells[5].Value = myObj.Tg009;   // 預計結案
						dgv.Rows[i].Cells[6].Value = myObj.Tg010;   // 量產地
						dgv.Rows[i].Cells[7].Value = myObj.Tg011;   // 類型
						dgv.Rows[i].Cells[8].Value = myObj.Tg012;   // 耐用年限

						dgv.Rows[i].Cells[10].Value = myObj.Tg000;   // 唯一序號
					}
				}
            }
			catch (Exception ex)
			{
				_log.Debug("ANBTG_Model的SetDGV方式發生錯誤：" + ex.Message);
				MessageBox.Show("ANBTG_Model的SetDGV方式發生錯誤：" + ex.Message);
			}

			return dgv;
		}
	}
}
