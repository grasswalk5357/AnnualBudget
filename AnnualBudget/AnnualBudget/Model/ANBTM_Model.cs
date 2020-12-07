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
    class ANBTM_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		public static bool UpdateToANBTM(List<Object> list)
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
					foreach (ANBTM myObj in list)
					{
						SQL_1.Clear();
						SQL_2.Clear();

						SQL_1.AppendFormat("IF EXISTS (SELECT TM000 FROM dbo.ANBTM WHERE TM001 = {0} AND TM017 = '{1}' ) " +
											"UPDATE dbo.ANBTM SET FLAG = FLAG + 1, MODIFIER = '{2}', MODI_DATE = '{3}', " +
											"TM003 = {4}, TM004 = {5}, TM005 = {6}, TM006 = {7}, TM007 = {8}, " +
											"TM008 = {9}, TM009 = {10}, TM010 = {11}, TM011 = {12}, TM012 = {13}, " +
											"TM013 = {14}, TM014 = {15}, TM015 = {16}, TM016 = {17}, TM017 = '{18}', " +
											"TM018 = '{19}' " +
											"WHERE TM001 = {20} AND TM017 = '{21}'" +
											"ELSE ",
											myObj.Tm001, myObj.Tm017, myObj.Modifier, myObj.Modi_date,
											myObj.Tm003, myObj.Tm004, myObj.Tm005, myObj.Tm006, myObj.Tm007, 
											myObj.Tm008, myObj.Tm009, myObj.Tm010, myObj.Tm011, myObj.Tm012, 
											myObj.Tm013, myObj.Tm014, myObj.Tm015, myObj.Tm016, myObj.Tm017,
											myObj.Tm018, myObj.Tm001, myObj.Tm017
											);


						SQL_2.AppendFormat("INSERT INTO dbo.ANBTM " +
							"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, " +
							"TM001, TM002, TM003, TM004, TM005, " +
							"TM006, TM007, TM008, TM009, TM010, " +
							"TM011, TM012, TM013, TM014, TM015, " +
							"TM016, TM017, TM018 " +
							") " +
							"VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, " +
							"'{7}', '{8}', {9}, {10}, {11}, " +
							"{12}, {13}, {14}, {15}, {16}, " +
							"{17}, {18}, {19}, {20}, {21}, " +
							"{22}, '{23}', '{24}' " +
							" ) ",
							myObj.Company, myObj.Creator, myObj.Usr_group, myObj.Create_date, "", "", myObj.Flag,
							myObj.Tm001, myObj.Tm002, myObj.Tm003, myObj.Tm004, myObj.Tm005,
							myObj.Tm006, myObj.Tm007, myObj.Tm008, myObj.Tm009, myObj.Tm010,
							myObj.Tm011, myObj.Tm012, myObj.Tm013, myObj.Tm014, myObj.Tm015,
							myObj.Tm016, myObj.Tm017, myObj.Tm018);

						SQL_1 = SQL_1.Append(SQL_2);


						//command = new SqlCommand(SQL_1.ToString(), conn);
						command.CommandText = SQL_1.ToString();
						command.ExecuteNonQuery();
					}
					trans.Commit();

				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTM」時發生錯誤：" + e.Message);
					MessageBox.Show("更新至「ANBTM」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「ANBTM」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTM」時發生錯誤的訊息：" + e2.Message);
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


		public static List<Object> Load_ANBTM(string FeeType, string year)
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
					SQL.AppendFormat("SELECT TM000, TM001, TM002, TM003, TM004, " +
									 "TM005, TM006, TM007, TM008, TM009, " +
									 "TM010, TM011, TM012, TM013, TM014, " +
									 "TM015, TM016, TM017, TM018 " +									 
									 "FROM dbo.ANBTM " +
									 "WHERE TM017 = '{0}' ", FeeType);

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
							ANBTM myObject = new ANBTM();
							myObject.Tm000 = Convert.ToDecimal(dt.Rows[i]["TM000"]);   // 流水號
							myObject.Tm001 = dt.Rows[i]["TM001"].ToString();           // 單別代號
							myObject.Tm002 = dt.Rows[i]["TM002"].ToString();           // 部門代號
							myObject.Tm003 = Convert.ToDecimal(dt.Rows[i]["TM003"].ToString());           // 年度
							myObject.Tm004 = Convert.ToDecimal(dt.Rows[i]["TM004"].ToString());
							myObject.Tm005 = Convert.ToDecimal(dt.Rows[i]["TM005"].ToString());
							myObject.Tm006 = Convert.ToDecimal(dt.Rows[i]["TM006"].ToString());
							myObject.Tm007 = Convert.ToDecimal(dt.Rows[i]["TM007"].ToString());
							myObject.Tm008 = Convert.ToDecimal(dt.Rows[i]["TM008"].ToString());
							myObject.Tm009 = Convert.ToDecimal(dt.Rows[i]["TM009"].ToString());
							myObject.Tm010 = Convert.ToDecimal(dt.Rows[i]["TM010"].ToString());
							myObject.Tm011 = Convert.ToDecimal(dt.Rows[i]["TM011"].ToString());
							myObject.Tm012 = Convert.ToDecimal(dt.Rows[i]["TM012"]);
							myObject.Tm013 = Convert.ToDecimal(dt.Rows[i]["TM013"].ToString());
							myObject.Tm014 = Convert.ToDecimal(dt.Rows[i]["TM014"].ToString());
							myObject.Tm015 = Convert.ToDecimal(dt.Rows[i]["TM015"].ToString());
							myObject.Tm016 = Convert.ToDecimal(dt.Rows[i]["TM016"].ToString());
							myObject.Tm017 = dt.Rows[i]["TM017"].ToString();
							myObject.Tm018 = dt.Rows[i]["TM018"].ToString();

							list.Add(myObject);
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("載入[ANBTM]資料時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return list;
			}
		}

		public static List<Object> GetOtherFee(string deptNo, string FeeCode1, string FeeCode2) {

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
					SQL.AppendFormat("SELECT TM001, TM002, " +
						"SUM(TM003) AS TM003, SUM(TM004) AS TM004, SUM(TM005) AS TM005, " +
						"SUM(TM006) AS TM006, SUM(TM007) AS TM007, SUM(TM008) AS TM008, " +
						"SUM(TM009) AS TM009, SUM(TM010) AS TM010, SUM(TM011) AS TM011, " +
						"SUM(TM012) AS TM012, SUM(TM013) AS TM013, SUM(TM014) AS TM014, " +
						"SUM(TM015) AS TM015, SUM(TM016) AS TM016, '{0}' AS TM017 " +
						"FROM dbo.ANBTM ", FeeCode1);

					if (String.IsNullOrEmpty(FeeCode2))
						SQL.AppendFormat("WHERE TM017 = '{0}' ", FeeCode1);
					else
						SQL.AppendFormat("WHERE (TM017 = '{0}' OR TM017 = '{1}') ", FeeCode1, FeeCode2);

					SQL.AppendFormat("AND TM001 = '{0}' " +
						"GROUP BY TM001, TM002 " +
						"ORDER BY TM001", deptNo);


					

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
							ANBTM myObject = new ANBTM();
							//myObject.Tm000 = Convert.ToDecimal(dt.Rows[i]["TM000"]);   // 流水號
							myObject.Tm001 = dt.Rows[i]["TM001"].ToString();           // 部門代號
							myObject.Tm002 = dt.Rows[i]["TM002"].ToString();           // 部門名稱
							myObject.Tm003 = Convert.ToDecimal(dt.Rows[i]["TM003"].ToString());  // 費率
							myObject.Tm004 = Convert.ToDecimal(dt.Rows[i]["TM004"].ToString());  // 1月
							myObject.Tm005 = Convert.ToDecimal(dt.Rows[i]["TM005"].ToString());  // 2月
							myObject.Tm006 = Convert.ToDecimal(dt.Rows[i]["TM006"].ToString());  // 3月
							myObject.Tm007 = Convert.ToDecimal(dt.Rows[i]["TM007"].ToString());  // 4月
							myObject.Tm008 = Convert.ToDecimal(dt.Rows[i]["TM008"].ToString());  // 5月
							myObject.Tm009 = Convert.ToDecimal(dt.Rows[i]["TM009"].ToString());  // 6月
							myObject.Tm010 = Convert.ToDecimal(dt.Rows[i]["TM010"].ToString());  // 7月
							myObject.Tm011 = Convert.ToDecimal(dt.Rows[i]["TM011"].ToString());  // 8月
							myObject.Tm012 = Convert.ToDecimal(dt.Rows[i]["TM012"].ToString());  // 9月
							myObject.Tm013 = Convert.ToDecimal(dt.Rows[i]["TM013"].ToString());  // 10月
							myObject.Tm014 = Convert.ToDecimal(dt.Rows[i]["TM014"].ToString());  // 11月
							myObject.Tm015 = Convert.ToDecimal(dt.Rows[i]["TM015"].ToString());  // 12月
							myObject.Tm016 = Convert.ToDecimal(dt.Rows[i]["TM016"].ToString());  // 合計
							myObject.Tm017 = dt.Rows[i]["TM017"].ToString();           // 部門代號

							list.Add(myObject);
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("依部門載入[ANBTM]資料時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return list;
			}			
		}

	}
}
