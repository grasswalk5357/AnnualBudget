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
    class ANBTH_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// 新增或更新資料至ANBTH資料表
		/// </summary>
		/// <param name="list">存放ANBTH物件的List</param>
		/// <returns></returns>
		public static bool UpdateToANBTH(List<Object> list)
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
					
					for (int i = 0; i < list.Count; i++)
					{
						ANBTH myObj = (ANBTH)list[i];
						SQL_1.Clear();
						SQL_2.Clear();
						
						SQL_1.AppendFormat("IF EXISTS (SELECT TH000 FROM dbo.ANBTH WHERE TH000 = {0} ) " +
											"UPDATE dbo.ANBTH SET FLAG = FLAG + 1, MODIFIER = '{1}', MODI_DATE = '{2}', " +
											"TH004 = '{3}', TH005 = '{4}', TH006 = '{5}', TH007 = {6}, TH008 = {7}, " +
											"TH009 = {8}, TH010 = {9}, TH011 = {10}, TH012 = {11}, TH013 = {12}, " +
											"TH014 = {13}, TH015 = {14}, TH016 = {15}, TH017 = {16}, TH018 = {17}, " +
											"TH019 = {18}, TH020 = {19}, TH021 = {20}, TH022 = {21}, TH023 = '{22}' " +
											"WHERE TH000 = {23}  ELSE ",
											myObj.Th000, myObj.Modifier, myObj.Modi_date,
											myObj.Th004, myObj.Th005, myObj.Th006, 
											myObj.MonthlyData[0], myObj.MonthlyData[1], myObj.MonthlyData[2], myObj.MonthlyData[3],
											myObj.MonthlyData[4], myObj.MonthlyData[5], myObj.MonthlyData[6], myObj.MonthlyData[7], 
											myObj.MonthlyData[8], myObj.MonthlyData[9], myObj.MonthlyData[10], myObj.MonthlyData[11], 
											myObj.Th019, myObj.Th020, myObj.Th021, myObj.Th022, myObj.Th023, 
											myObj.Th000);
						

						


						SQL_2.AppendFormat("INSERT INTO dbo.ANBTH " +
							"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, " +
							"TH001, TH002, TH003, TH004, TH005, TH006, " +
							"TH007, TH008, TH009, TH010, TH011, TH012, " +
							"TH013, TH014, TH015, TH016, TH017, TH018, " +
							"TH019, TH020, TH021, TH022, TH023) " +
							"VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, " +
							"'{7}', '{8}', '{9}', '{10}', '{11}', " +
							"'{12}', {13}, {14}, {15}, {16}, " +
							"{17}, {18}, {19}, {20}, {21}, " +
							"{22}, {23}, {24}, {25}, {26}, " +
							"{27}, {28}, '{29}' " +
							" ) ",
							myObj.Company, myObj.Creator, myObj.Usr_group, myObj.Create_date, "", "", myObj.Flag,
							myObj.Th001, myObj.Th002, myObj.Th003, myObj.Th004, myObj.Th005,
							myObj.Th006, myObj.MonthlyData[0], myObj.MonthlyData[1], myObj.MonthlyData[2], myObj.MonthlyData[3],
							myObj.MonthlyData[4], myObj.MonthlyData[5], myObj.MonthlyData[6], myObj.MonthlyData[7], myObj.MonthlyData[8],
							myObj.MonthlyData[9], myObj.MonthlyData[10], myObj.MonthlyData[11], myObj.Th019, myObj.Th020, 
							myObj.Th021, myObj.Th022, myObj.Th023);

						SQL_1 = SQL_1.Append(SQL_2);


						//command = new SqlCommand(SQL_1.ToString(), conn);
						command.CommandText = SQL_1.ToString();
						command.ExecuteNonQuery();
					}
					trans.Commit();

				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTH」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「ANBTH」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTH」時發生錯誤的訊息：" + e2.Message);
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


		public static Dictionary<string, object> LoadData(Dictionary<string, Object> list)
		{
			Dictionary<string, Object> keyValues = new Dictionary<string, object>();
			StringBuilder SQL = new StringBuilder();
			SqlDataAdapter adapter;
			DataTable dt = new DataTable();			

			SqlCommand command = null;
			EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
			SqlConnectionStringBuilder sqlsb;

			sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["AB_ConnString"].ConnectionString);

			sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

			using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
			{

				
				try
				{
					for (int x = 0; x < list.Count; x++) 					
					{
						Object obj = list.Values.ElementAt(x);
						ANBTG anbtg = (ANBTG)obj;

						SQL.Clear();
						dt.Clear();
						SQL.AppendFormat("SELECT TH000, TH001, TH002, TH003, TH004, " +
												"TH005, TH006, TH007, TH008, TH009, " +
												"TH010, TH011, TH012, TH013, TH014, " +
												"TH015, TH016, TH017, TH018, TH019, " +
												"TH020, TH021, TH022 " +
												"FROM dbo.ANBTH " +
										 "WHERE TH001 = '{0}' AND TH002 = '{1}' " +
										 "AND TH003 = '{2}' AND TH004 = '{3}' AND TH023 <> 'Y'", anbtg.Tg001, anbtg.Tg002, anbtg.Tg003, anbtg.Tg004);

						if (conn.State != ConnectionState.Open)
							conn.Open();
						command = new SqlCommand(SQL.ToString(), conn);

						adapter = new SqlDataAdapter(command);
						
						adapter.Fill(dt);

						if (dt.Rows.Count > 0)
						{

							//list = new List<object>();
							//for (int i = 0; i < dt.Rows.Count; i++)
							for (int i = 0; i < 9; i++)
							{
								ANBTH myObject = new ANBTH();
								myObject.Th000 = Convert.ToDecimal(dt.Rows[i]["Th000"]);   // 流水號
								myObject.Th001 = dt.Rows[i]["Th001"].ToString();           // 單別代號
								myObject.Th002 = dt.Rows[i]["Th002"].ToString();           // 部門代號
								myObject.Th003 = dt.Rows[i]["Th003"].ToString();           // 年度
								myObject.Th004 = dt.Rows[i]["Th004"].ToString();
								myObject.Th005 = dt.Rows[i]["Th005"].ToString();
								myObject.Th006 = dt.Rows[i]["Th006"].ToString();

								myObject.MonthlyData[0] = Convert.ToDecimal(dt.Rows[i]["Th007"]);
								myObject.MonthlyData[1] = Convert.ToDecimal(dt.Rows[i]["Th008"]);
								myObject.MonthlyData[2] = Convert.ToDecimal(dt.Rows[i]["Th009"]);
								myObject.MonthlyData[3] = Convert.ToDecimal(dt.Rows[i]["Th010"]);
								myObject.MonthlyData[4] = Convert.ToDecimal(dt.Rows[i]["Th011"]);
								myObject.MonthlyData[5] = Convert.ToDecimal(dt.Rows[i]["Th012"]);
								myObject.MonthlyData[6] = Convert.ToDecimal(dt.Rows[i]["Th013"]);
								myObject.MonthlyData[7] = Convert.ToDecimal(dt.Rows[i]["Th014"]);
								myObject.MonthlyData[8] = Convert.ToDecimal(dt.Rows[i]["Th015"]);
								myObject.MonthlyData[9] = Convert.ToDecimal(dt.Rows[i]["Th016"]);
								myObject.MonthlyData[10] = Convert.ToDecimal(dt.Rows[i]["Th017"]);
								myObject.MonthlyData[11] = Convert.ToDecimal(dt.Rows[i]["Th018"]);

								myObject.Th019 = Convert.ToDecimal(dt.Rows[i]["Th019"]);
								myObject.Th020 = Convert.ToDecimal(dt.Rows[i]["Th020"]);
								myObject.Th021 = Convert.ToDecimal(dt.Rows[i]["Th021"]);
								myObject.Th022 = Convert.ToDecimal(dt.Rows[i]["Th022"]);

								//anbtg.ANBTH_List.Add(myObject);
								anbtg.ANBTH_List[i] = myObject;
								
							}
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("從ANBTH取資料時發生錯誤：" + e.Message);
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
		/// <param name="dgv_PRJ">研發專案彙總表_單頭DataGridView</param>
		/// <param name="dgv_PRJ_Content">研發專案彙總表_單身DataGridView</param>
		/// <param name="keyValues">配對單頭與單身資料的Dictionary物件</param>
		/// <returns></returns>
		public static DataGridView SetDGV(DataGridView dgv_PRJ, DataGridView dgv_PRJ_Content, Dictionary<string, Object> keyValues) {
			if (dgv_PRJ.SelectedCells != null && dgv_PRJ.SelectedCells.Count > 0)
			{
				string prjid = dgv_PRJ.Rows[dgv_PRJ.SelectedCells[0].RowIndex].Cells[0].Value.ToString();
				
				ANBTG rdp = null;

				if (keyValues.Count > 0)
				{
					if (keyValues.ContainsKey(prjid))						
						rdp = (ANBTG)keyValues[prjid];
					

					if (rdp != null)
					{
						//for (int i = 0; i < rdp.RDP_Content_List.Count; i++)
						for (int i = 0; i < rdp.ANBTH_List.Count; i++)
						{
							ANBTH anbth = (ANBTH)rdp.ANBTH_List[i];
							for (int j = 0; j < dgv_PRJ_Content.ColumnCount; j++)    // 所有的欄位，一直到chk_box之前
							{
								if (j != 2 || j != 4) // 單身的第3(rowIndex = 2)跟第5(rowIndex = 4)個項目暫時不使用
								{ 
									if (j < 12) // 月份的欄位
									{

										if (anbth.MonthlyData[j] > 0)
											dgv_PRJ_Content.Rows[i].Cells[j].Value = anbth.MonthlyData[j];
										else
											dgv_PRJ_Content.Rows[i].Cells[j].Value = null;
									}
								}
							}

							dgv_PRJ_Content.Rows[i].Cells[12].Value = anbth.Th019;
							dgv_PRJ_Content.Rows[i].Cells[13].Value = anbth.Th020;
							dgv_PRJ_Content.Rows[i].Cells[14].Value = anbth.Th021;
							dgv_PRJ_Content.Rows[i].Cells[15].Value = anbth.Th022;
							dgv_PRJ_Content.Rows[i].Cells[16].Value = anbth.Th000;
			
						}
					}
					else
					{
						for (int i = 0; i < dgv_PRJ_Content.RowCount; i++)
						{
							for (int j = 0; j < dgv_PRJ_Content.ColumnCount; j++)
							{
								dgv_PRJ_Content.Rows[i].Cells[j].Value = null;
							}
						}
					}
				}
			}
			return dgv_PRJ_Content;
		}
	}
}
