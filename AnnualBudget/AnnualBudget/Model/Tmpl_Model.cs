using AnnualBudget.BOs;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.Model
{
    class Tmpl_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		public static DataTable GetDept_Table()
		{			
			StringBuilder selectString = new StringBuilder();
			

			EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
			SqlConnectionStringBuilder sqlsb;

			sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["myConnString"].ConnectionString);

			sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

			using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
			{

				SqlCommand command = null;
				SqlDataAdapter adapter = null;				
				DataTable table = new DataTable();

				try
				{
					selectString.Clear();

					
					selectString.AppendFormat("SELECT ME001 AS '部門代號', ME002 AS '部門名稱' FROM dbo.CMSME WHERE Len(ME001) = 3 AND CHARINDEX('A', ME001) <= 0");

					conn.Open();
					command = new SqlCommand(selectString.ToString(), conn);
					adapter = new SqlDataAdapter(command);

					adapter.Fill(table);

					for (int i = 0; i < table.Rows.Count; i++) {
						for (int j = 0; j < table.Columns.Count; j++) { 
							table.Rows[i][j] = table.Rows[i][j].ToString().Trim();							
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("取[部門代號]時發生錯誤：" + e.Message);
				}
				finally
				{
					if (adapter != null) adapter.Dispose();
					if (command != null) command.Dispose();
					//if (ds != null) ds.Dispose();
				}

				return table;
			}
		}

		public static DataTable GetDept_Tmpl_Ref_Table(string tmplID, string ver)
		{
			StringBuilder selectString = new StringBuilder();
			EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
			SqlConnectionStringBuilder sqlsb;

			sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["AB_ConnString"].ConnectionString);

			sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

			using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
			{
				SqlCommand command = null;
				SqlDataAdapter adapter = null;
				DataTable table = new DataTable();

				try
				{
					selectString.Clear();

					selectString.AppendFormat("SELECT TK002 AS '部門代號', '' AS '標記刪除' FROM dbo.ANBTK WHERE TK004 = '{0}' AND TK005 = '{1}' ORDER BY '部門代號'", tmplID, ver);

					conn.Open();
					command = new SqlCommand(selectString.ToString(), conn);
					adapter = new SqlDataAdapter(command);

					adapter.Fill(table);
				}
				catch (Exception e)
				{
					_log.Debug("取[部門代號與會科樣版對應]時發生錯誤：" + e.Message);
				}
				finally
				{
					if (adapter != null) adapter.Dispose();
					if (command != null) command.Dispose();
					//if (ds != null) ds.Dispose();
				}

				return table;
			}
		}


		public static DataTable GetTmplID()
		{
			StringBuilder selectString = new StringBuilder();
			EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
			SqlConnectionStringBuilder sqlsb;

			sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["AB_ConnString"].ConnectionString);

			sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

			using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
			{
				SqlCommand command = null;
				SqlDataAdapter adapter = null;
				DataTable table = new DataTable();

				try
				{
					selectString.Clear();

					selectString.AppendFormat("SELECT Distinct TL002, TL009 FROM dbo.ANBTL");

					conn.Open();
					command = new SqlCommand(selectString.ToString(), conn);
					adapter = new SqlDataAdapter(command);

					adapter.Fill(table);
				}
				catch (Exception e)
				{
					_log.Debug("取[部門代號]時發生錯誤：" + e.Message);
				}
				finally
				{
					if (adapter != null) adapter.Dispose();
					if (command != null) command.Dispose();
					//if (ds != null) ds.Dispose();
				}

				return table;
			}
		}

		public static DataTable GetTmplVer(string TmplID)
		{
			StringBuilder selectString = new StringBuilder();
			EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
			SqlConnectionStringBuilder sqlsb;

			sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["AB_ConnString"].ConnectionString);

			sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

			using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
			{
				SqlCommand command = null;
				SqlDataAdapter adapter = null;
				DataTable table = new DataTable();

				try
				{
					selectString.Clear();
					selectString.AppendFormat("SELECT DISTINCT TL003 FROM dbo.ANBTL WHERE TL002 = '{0}'", TmplID);

					conn.Open();
					command = new SqlCommand(selectString.ToString(), conn);					
					adapter = new SqlDataAdapter(command);

					adapter.Fill(table);
				}
				catch (Exception e)
				{
					_log.Debug("取[樣版版號]時發生錯誤：" + e.Message);
				}
				finally
				{
					if (adapter != null) adapter.Dispose();
					if (command != null) command.Dispose();					
				}

				return table;
			}
		}

		/// <summary>
		/// 取得樣版內容
		/// </summary>
		/// <param name="tmplID">樣版編號</param>
		/// <param name="ver">樣版版本號</param>
		/// <returns></returns>
		public static DataTable GetTmplContent(string tmplID, string ver) {
			StringBuilder selectString = new StringBuilder();
			EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
			SqlConnectionStringBuilder sqlsb;

			sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["AB_ConnString"].ConnectionString);

			sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

			using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
			{

				SqlCommand command = null;
				SqlDataAdapter adapter = null;
				DataTable table = new DataTable();

				try
				{
					selectString.Clear();
					
					selectString.AppendFormat("SELECT TL004 AS '序號', TL005 AS '會科編號', TL006 AS '會科名稱', TL007 AS '會科說明', ISNULL(TL010, '') AS TL010, ISNULL(TL011, '') AS TL011 " +
												"FROM dbo.ANBTL " +
												"WHERE TL002 = '{0}' AND TL003 = '{1}' ORDER BY '序號'", tmplID, ver);
					

			

					conn.Open();
					command = new SqlCommand(selectString.ToString(), conn);
					adapter = new SqlDataAdapter(command);

					adapter.Fill(table);
				}
				catch (Exception e)
				{
					_log.Debug("取[樣版內容]時發生錯誤：" + e.Message);
				}
				finally
				{
					if (adapter != null) adapter.Dispose();
					if (command != null) command.Dispose();
				}

				return table;
			}
		}

		/// <summary>
		/// 透過部門編號取得相對應的樣版內容
		/// </summary>
		/// <param name="DeptNo">部門編號</param>
		/// <returns></returns>
		public static DataTable GetTmplContentByDept(string DeptNo)
		{
			StringBuilder selectString = new StringBuilder();
			EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
			SqlConnectionStringBuilder sqlsb;

			sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["AB_ConnString"].ConnectionString);

			sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

			using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
			{

				SqlCommand command = null;
				SqlDataAdapter adapter = null;
				DataTable table = new DataTable();

				try
				{
					selectString.Clear();
					

					selectString.AppendFormat("SELECT TL004 AS '序號', " +
						  					  "TL005 AS '會科編號', " +
											  "TL006 AS '會科名稱', " +
											  "TL007 AS '會科說明', " +
											  "ISNULL(TL010, '') AS TL010, " +
											  "ISNULL(TL011, '') AS TL011, " +
											  "TL002, TL003, " +
											  "TK002 AS '部門', " +
											  "TL001 AS '流水號', " +
											  "TL012 " +
											  "FROM dbo.ANBTL TL JOIN dbo.ANBTK TK ON TL.TL002 = TK.TK004 " +
											  "WHERE TK002 = '{0}' " +
											  "ORDER BY '序號'", DeptNo);

					conn.Open();
					command = new SqlCommand(selectString.ToString(), conn);
					adapter = new SqlDataAdapter(command);

					adapter.Fill(table);
					
				}
				catch (Exception e)
				{
					_log.Debug("[依部門取得樣版內容]時發生錯誤：" + e.Message);
				}
				finally
				{
					if (adapter != null) adapter.Dispose();
					if (command != null) command.Dispose();
				}

				return table;
			}
		}

		public static bool Update_Dept_Tmpl_Ref(List<ANBTK> TK_list, string Year, UserInfo userInfo) {
			
			bool result = true;
			StringBuilder SQL_1 = new StringBuilder();
			StringBuilder SQL_2 = new StringBuilder();			
			string today = DateTime.Now.ToString("yyyyMMdd");

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
					foreach (ANBTK anbtk in TK_list) { 
						SQL_1.Clear();
						SQL_2.Clear();
						
						SQL_1.AppendFormat("IF EXISTS (SELECT TK001 FROM dbo.ANBTK WHERE TK002 = '{0}' AND TK003 = '{1}') " +
												"UPDATE dbo.ANBTK SET MODIFIER = '{2}', MODI_DATE = '{3}', FLAG = FLAG + 1, TK004 = '{4}', TK005 = '{5}' " +
												"WHERE TK002 = '{6}' AND TK003 = '{7}' ELSE ", anbtk.Tk002, Year, userInfo.DeptNo, today, anbtk.Tk004, anbtk.Tk005, anbtk.Tk002, Year);

						SQL_2.AppendFormat("INSERT INTO dbo.ANBTK " +
							"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, FLAG, TK002, TK003, TK004, TK005) " +
							"VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, " +
							"'{7}', '{8}') ",
							anbtk.Company, anbtk.Creator, userInfo.DeptNo, anbtk.Create_date, anbtk.Flag,
							anbtk.Tk002, Year, anbtk.Tk004, anbtk.Tk005);

						SQL_1 = SQL_1.Append(SQL_2);


						//command = new SqlCommand(SQL_1.ToString(), conn);
						command.CommandText = SQL_1.ToString();
						command.ExecuteNonQuery();
					}
					trans.Commit();
				}
				catch (Exception e)
				{
					_log.Debug("更新至「部門與模版對應表」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「部門與模版對應表」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「部門與模版對應表」時發生錯誤的訊息：" + e2.Message);
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
	}
}
