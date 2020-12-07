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
    class ANBTO_Model 
    {
		private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		public static List<Object> LoadData()
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
					SQL.AppendFormat("SELECT TO001, TO002, TO003, TO004, TO005 " +
						"FROM dbo.ANBTO " +
						"ORDER BY TO005, TO001");

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
							ANBTO anbto = new ANBTO();
							
							anbto.To001 = Convert.ToDecimal(dt.Rows[i]["TO001"].ToString());
							anbto.To002 = dt.Rows[i]["TO002"].ToString();
							anbto.To003 = dt.Rows[i]["TO003"].ToString();
							anbto.To004 = dt.Rows[i]["TO004"].ToString();
							anbto.To005 = dt.Rows[i]["TO005"].ToString();

							list.Add(anbto);
						}
					}
				}
				catch (Exception e)
				{
					_log.Debug("載入「ANBTO」時發生錯誤：" + e.Message);
					MessageBox.Show("載入「ANBTO」時發生錯誤：" + e.Message);
				}
				finally
				{
					if (command != null) command.Dispose();
				}

				return list;
			}
		}


		public static List<Object> SaveMatrixData(DataGridView dgv, string userId, string deptNo)
		{
			List<Object> list = new List<Object>();
			ANBTO myObject = null;
			try
			{
				for (int i = 0; i < dgv.RowCount; i++)
				{
					if (dgv.Rows[i].Cells["dgv_AuthMgmt_IsEdited"].Value != null) 
					{
						if ("V".Equals(dgv.Rows[i].Cells["dgv_AuthMgmt_IsEdited"].Value.ToString()))	// 只有有編輯過的資料，才要進行儲存
						{
							myObject = new ANBTO();
							myObject.Creator = userId;
							myObject.Usr_group = deptNo;
							myObject.Create_date = DateTime.Now.ToString("yyyyMMdd");
							myObject.Modifier = userId;
							myObject.Modi_date = DateTime.Now.ToString("yyyyMMdd");
							myObject.DeptNo = deptNo;

							myObject.To001 = Convert.ToDecimal(dgv.Rows[i].Cells["dgv_AuthMgmt_TO001"].Value);

							if (dgv.Rows[i].Cells["dgv_AuthMgmt_TO002"].Value != null)
								myObject.To002 = dgv.Rows[i].Cells["dgv_AuthMgmt_TO002"].Value.ToString();

							if (dgv.Rows[i].Cells["dgv_AuthMgmt_TO003"].Value != null)
								myObject.To003 = dgv.Rows[i].Cells["dgv_AuthMgmt_TO003"].Value.ToString();

							if (dgv.Rows[i].Cells["dgv_AuthMgmt_TO004"].Value != null)
								myObject.To004 = dgv.Rows[i].Cells["dgv_AuthMgmt_TO004"].Value.ToString();

							if (dgv.Rows[i].Cells["dgv_AuthMgmt_TO005"].Value != null)
								myObject.To005 = dgv.Rows[i].Cells["dgv_AuthMgmt_TO005"].Value.ToString();

							list.Add(myObject);
						}
					}
					
				}
			}
			catch (Exception e)
			{
				_log.Debug("儲存[人員權限管理]時發生錯誤：" + e.Message);
				MessageBox.Show("儲存[人員權限管理]時發生錯誤：" + e.Message);
			}

			return list;
		}


		public static bool UpdateToANBTO(List<Object> myList)
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
					foreach (ANBTO myObj in myList)
					{
						SQL_1.Clear();
						SQL_2.Clear();

						SQL_1.AppendFormat("IF EXISTS (SELECT TO001 FROM dbo.ANBTO WHERE TO001 = {0} ) " +
												"UPDATE dbo.ANBTO SET FLAG = FLAG + 1, MODIFIER = '{1}', MODI_DATE = '{2}', " +
												"TO002 = '{3}', TO003 = '{4}', TO004 = '{5}', TO005 = '{6}' " +
												"WHERE TO001 = {7} ELSE ",
												myObj.To001, myObj.Modifier, myObj.Modi_date,
												myObj.To002, myObj.To003, myObj.To004, myObj.To005,
												myObj.To001);

						SQL_2.AppendFormat("INSERT INTO dbo.ANBTO " +
							"(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, " +
							"TO002, TO003, TO004, TO005) " +
							"VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, " +
							"'{7}', '{8}', '{9}', '{10}' ) ",
							myObj.Company, myObj.Creator, myObj.Usr_group, myObj.Create_date, "", "", myObj.Flag,
							myObj.To002, myObj.To003, myObj.To004, myObj.To005);

						SQL_1 = SQL_1.Append(SQL_2);

						command.CommandText = SQL_1.ToString();
						command.ExecuteNonQuery();

					}
					trans.Commit();
				}
				catch (Exception e)
				{
					_log.Debug("更新至「ANBTO」時發生錯誤：" + e.Message);
					try
					{
						trans.Rollback();
					}
					catch (Exception e2)
					{
						_log.Debug("更新至「ANBTO」時發生錯誤的RollBack Exp Type：" + e2.GetType());
						_log.Debug("更新至「ANBTO」時發生錯誤的訊息：" + e2.Message);
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
