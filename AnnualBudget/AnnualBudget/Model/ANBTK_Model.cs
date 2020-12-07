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
    class ANBTK_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


		/// <summary>
		/// 讀取登入者的部門、上級部門與會計科目樣版的對應資訊
		/// </summary>
		/// <param name="userinfo">登入者資訊物件</param>
		/// <param name="year">年份</param>
		/// <param name="List_Depts">儲存結果的List</param>
		/// <returns></returns>
		public static List<Object> LoadDepts(UserInfo userinfo, string year, List<Object> List_Depts) 
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

					if (userinfo.DeptNo.Equals("510"))
					{
						string head = userinfo.DeptNo.Substring(0, 2);
						SQL.AppendFormat("SELECT TK002 FROM dbo.ANBTK WHERE TK002 like '{0}%' AND TK003 = '{1}' ORDER BY TK002", head, year);
					}
					/*else if (userinfo.ParentDept != null)
						SQL.AppendFormat("SELECT TK002 FROM dbo.ANBTK WHERE TK002 = '{0}' AND TK003 = '{1}' ORDER BY TK002", userinfo.DeptNo, year);					
					else
					{
						if (searchMode == 2) { 
							string head = userinfo.DeptNo.First().ToString();
							SQL.AppendFormat("SELECT TK002 FROM dbo.ANBTK WHERE TK002 like '{0}%' AND TK003 = '{1}' ORDER BY TK002", head, year);
						}
					}*/

					else if (userinfo.Level == "3")  // 科級別
					{
						SQL.AppendFormat("SELECT TK002 FROM dbo.ANBTK WHERE TK002 = '{0}' AND TK003 = '{1}' ", userinfo.DeptNo, year);
						if (userinfo.ParentDept != null) {
							SQL.Append(" UNION ");
							SQL.AppendFormat("SELECT TK002 FROM dbo.ANBTK WHERE TK002 = '{0}' AND TK003 = '{1}' ", userinfo.ParentDept.DeptNo, year);
						}

						SQL.Append(" ORDER BY TK002 ");
					}
					else if (userinfo.Level == "2") // 部級別
					{
						SQL.AppendFormat("SELECT TK002 FROM dbo.ANBTK WHERE TK002 = '{0}' AND TK003 = '{1}' ", userinfo.DeptNo, year);
						if (userinfo.ParentDept != null)
						{
							SQL.Append(" union ");
							SQL.AppendFormat("SELECT TK002 FROM dbo.ANBTK WHERE TK002 = '{0}' AND TK003 = '{1}' ", userinfo.ParentDept.DeptNo, year);
						}
						SQL.Append(" ORDER BY TK002 ");
					}
					else if (userinfo.Level == "1") // 處級別
					{
						string head = userinfo.DeptNo.First().ToString();
						SQL.AppendFormat("SELECT TK002 FROM dbo.ANBTK WHERE TK002 like '{0}%' AND TK003 = '{1}' ORDER BY TK002", head, year);
					}

					/*else
					{
						if (searchMode == 2)
						{
							string head = userinfo.DeptNo.First().ToString();
							SQL.AppendFormat("SELECT TK002 FROM dbo.ANBTK WHERE TK002 like '{0}%' AND TK003 = '{1}' ORDER BY TK002", head, year);
						}
					}*/

					if (!String.IsNullOrEmpty(SQL.ToString())) { 
						if (conn.State != ConnectionState.Open)
							conn.Open();
						command = new SqlCommand(SQL.ToString(), conn);
										
						adapter = new SqlDataAdapter(command);

						adapter.Fill(dt);

						if (dt.Rows.Count > 0) {
							//List_Depts.Add(userinfo.DeptNo);
							for (int i = 0; i < dt.Rows.Count; i++) { 
								if (!List_Depts.Contains(dt.Rows[i]["TK002"].ToString()))
									List_Depts.Add(dt.Rows[i]["TK002"].ToString());
							}
						}
					}

					//if (userinfo.ParentDept != null) {
					//	LoadDepts(userinfo.ParentDept, year, List_Depts, 1);
					//}
					
				}
				catch (Exception e)
				{
					_log.Debug("讀取「ANBTK」時發生錯誤：" + e.Message);					
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return List_Depts;
			}
        }



		/// <summary>
		/// 取得所有要填年度預算的部門
		/// </summary>
		/// <param name="year"></param>
		/// <param name="List_Depts"></param>
		/// <returns></returns>
		public static DataTable LoadDepts(string year)
		{
			//List<Object> List = new List<object>();
			//ANBTK anbtk = null;

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
					SQL.AppendFormat("SELECT TK001, TK002, TK003, TK004, TK005 " +
						"FROM dbo.ANBTK " +
						"WHERE TK004 <> '' " +
						"AND TK003 = '{0}' ORDER BY TK002", year);

					if (conn.State != ConnectionState.Open)
						conn.Open();
					command = new SqlCommand(SQL.ToString(), conn);

					adapter = new SqlDataAdapter(command);

					adapter.Fill(dt);
					
				}
				catch (Exception e)
				{
					_log.Debug("讀取「ANBTK」時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return dt;
			}
		}


		public static DataTable LoadDepts()
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
					SQL.AppendFormat("SELECT TK002 AS '部門代號' FROM dbo.ANBTK WHERE TK006 = 'Y' order by TK002");

					if (conn.State != ConnectionState.Open)
						conn.Open();
					command = new SqlCommand(SQL.ToString(), conn);

					adapter = new SqlDataAdapter(command);

					adapter.Fill(dt);

				}
				catch (Exception e)
				{
					_log.Debug("執行「LoadDepts()」方法時發生錯誤：" + e.Message);
				    MessageBox.Show("執行「LoadDepts()」方法時發生錯誤：" + e.Message);
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
