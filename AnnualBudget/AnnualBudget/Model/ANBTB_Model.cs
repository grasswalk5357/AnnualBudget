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
    class ANBTB_Model
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// 新增或更新資料至ANBTB資料表
		/// </summary>
		/// <param name="anbtb"></param>
		/// <returns></returns>
		public static bool UpdateToANBTB(ANBTB anbtb)
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
					SQL_1.Clear();
					SQL_2.Clear();

					SQL_1.AppendFormat("IF EXISTS (SELECT TB001 FROM dbo.ANBTB WHERE TB002 = '{0}' AND TB003 = '{1}' AND TB004 = '{2}') " +
											"UPDATE dbo.ANBTB SET TB006 = '{3}', FLAG = FLAG + 1, MODIFIER = '{4}', MODI_DATE = '{5}' " +
											"WHERE TB002 = '{6}' AND TB003 = '{7}' AND TB004 = '{8}' ELSE ", 
											anbtb.Tb002, anbtb.Tb003, anbtb.Tb004, anbtb.Tb006, anbtb.Modifier, anbtb.Modi_date, anbtb.Tb002, anbtb.Tb003, anbtb.Tb004);

					SQL_2.AppendFormat("INSERT INTO dbo.ANBTB " +
						"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, FLAG, TB002, TB003, TB004, TB005, TB006) " +
						"VALUES('{0}', '{1}', '{2}', '{3}', {4}, '{5}', '{6}', " +
						"'{7}', '{8}', '{9}') ",
						anbtb.Company, anbtb.Creator, anbtb.Usr_group, anbtb.Create_date, anbtb.Flag,
						anbtb.Tb002, anbtb.Tb003, anbtb.Tb004, anbtb.Tb005, anbtb.Tb006);

					SQL_1 = SQL_1.Append(SQL_2);


					//command = new SqlCommand(SQL_1.ToString(), conn);
					command.CommandText = SQL_1.ToString();
					command.ExecuteNonQuery();
					trans.Commit();
					
				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTB」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「ANBTB」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTB」時發生錯誤的訊息：" + e2.Message);
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
