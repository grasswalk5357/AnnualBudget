using AnnualBudget.BOs;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.Model
{
    class ANBTJ_Model
    {
        private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// 取得勞、健保、伙食費及退休金等各項費率表
		/// </summary>
		/// <returns></returns>
		public static DataTable GetRateTable()
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
										
					selectString.AppendFormat("SELECT TJ002 AS '費率代號', TJ003 AS '費率名稱', TJ004 AS '費率', TJ005 AS '備註', '' AS '已編輯' " +
											  "FROM dbo.ANBTJ");
					conn.Open();
					command = new SqlCommand(selectString.ToString(), conn);
					adapter = new SqlDataAdapter(command);

					adapter.Fill(table);
				}
				catch (Exception e)
				{
					_log.Debug("取[相關費率表]時發生錯誤：" + e.Message);
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
		/// 取得勞、健保、伙食費或退休金費率
		/// </summary>
		/// <param name="RateTypeCode">R001:勞保, R002:健保, R003: 伙食費, R004:退休金提撥費率</param>
		/// <returns></returns>
		public static decimal GetRate(string RateTypeCode)
		{
			decimal result = 0;
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

					selectString.AppendFormat("SELECT TJ004 FROM dbo.ANBTJ WHERE TJ002 = '{0}'", RateTypeCode);

					conn.Open();
					command = new SqlCommand(selectString.ToString(), conn);
					adapter = new SqlDataAdapter(command);

					adapter.Fill(table);
					
					if (table != null) {
						result = Convert.ToDecimal(table.Rows[0][0].ToString());
					}
				}
				catch (Exception e)
				{
					_log.Debug("取[單一費率值]時發生錯誤：" + e.Message);
				}
				finally
				{
					if (adapter != null) adapter.Dispose();
					if (command != null) command.Dispose();
				}

				return result;
			}
		}

		/// <summary>
		/// 新增或更新資料至ANBTJ資料表
		/// </summary>
		/// <param name="list">存放ANBTJ物件的List</param>
		/// <param name="UserID">登入者工號</param>
		/// <returns></returns>
		public static bool Update_Rate(List<ANBTJ> list, string UserID)
		{
			string today = DateTime.Now.ToString("yyyyMMdd");
			bool result = true;
			StringBuilder SQL_1 = new StringBuilder();
			StringBuilder SQL_2 = new StringBuilder();
			EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
			SqlConnectionStringBuilder sqlsb;

			sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["AB_ConnString"].ConnectionString);

			sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

			using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
			{
				SqlCommand command = null;
				SqlTransaction trans = null;
				try
				{
					
					if (list != null && list.Count > 0) {

						if (conn.State != ConnectionState.Open)
							conn.Open();

						command = conn.CreateCommand();
						trans = conn.BeginTransaction(IsolationLevel.ReadCommitted);
						command.Connection = conn;
						command.Transaction = trans;

						foreach (ANBTJ obj in list)
						{
							SQL_1.Clear();
							SQL_2.Clear();

							SQL_1.AppendFormat("IF EXISTS (SELECT TJ001 FROM dbo.ANBTJ WHERE TJ002 = '{0}') " +
											   "UPDATE dbo.ANBTJ SET MODIFIER = '{1}', MODI_DATE = '{2}', FLAG = FLAG + 1, TJ003 = '{3}', TJ004 = {4}, TJ005 = '{5}' " +
											   "WHERE TJ002 = '{6}' ELSE ", obj.Tj002, UserID, today, obj.Tj003, obj.Tj004, obj.Tj005, obj.Tj002);

							SQL_2.AppendFormat("INSERT INTO dbo.ANBTJ " +
											   "(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, FLAG, TJ002, TJ003, TJ004, TJ005) " +
											   "VALUES('{0}', '{1}', '{2}', '{3}', {4}, '{5}', '{6}', " +
											   "{7}, '{8}') ", 
											   obj.Company, obj.Creator, obj.Usr_group, obj.Create_date, obj.Flag, 
											   obj.Tj002, obj.Tj003, obj.Tj004, obj.Tj005);

							SQL_1 = SQL_1.Append(SQL_2);


							//command = new SqlCommand(SQL_1.ToString(), conn);
							command.CommandText = SQL_1.ToString();
							command.ExecuteNonQuery();
						}
						trans.Commit();
					}
				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTJ」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「ANBTJ」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTJ」時發生錯誤的訊息：" + e2.Message);
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
