using AnnualBudget.BOs;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.Model
{
    
    class Auth_Model
    {
        private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string AP_Server = ConfigurationManager.AppSettings["AP_Server"];

        /// <summary>
        /// 驗證AD網域帳號、密碼是否合法
        /// </summary>
        /// <param name="empID">AD帳號</param>
        /// <param name="pwd">AD密碼</param>
        /// <returns>TRUE :成功  FALSE:失敗</returns>
        public Boolean ValidateKrtcAD(String empID, String pwd)
        {
            
            DirectoryEntry de = new DirectoryEntry(AP_Server, empID, pwd);

            try
            {
                IsPasswordUseful(empID);
                Object obj = de.NativeObject; //不出現錯錯誤表示帳密成功

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                de.Close();
            }
        }

        /// <summary>
        /// 檢查密碼是否有效
        /// </summary>
        /// <param name="empID">AD帳號</param>
        private void IsPasswordUseful(String empID)
        {
            DirectoryEntry de = new DirectoryEntry(AP_Server);
            DirectorySearcher ds = new DirectorySearcher(de);
            ds.Filter = "(SAMAccountName=" + empID + ")";

            try
            {
                SearchResult sr = ds.FindOne();
                if (null == sr)
                {
                    throw new Exception();
                }
                else
                {
                    DirectoryEntry childDe = new DirectoryEntry(sr.Path);

                    DateTime dtExpireDate = Convert.ToDateTime(childDe.InvokeGet("PasswordExpirationDate").ToString());

                    //userAccountControl=66048表示密碼永久有效
                    int intuserAccountControl = (int)childDe.Properties["userAccountControl"].Value;

                    //密碼到期時間小於當下時間 , 密碼不是永久有效
                    if (DateTime.Compare(dtExpireDate, DateTime.Now) < 0 && 66048 != intuserAccountControl)
                    {
                        throw new Exception();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                ds.Dispose();
                de.Close();
            }
        }


        /// <summary>
        /// 透過員工代號取得姓名、部門名稱、部門代號
        /// </summary>
        /// <param name="userId"></param>
        /// <returns></returns>
        public static UserInfo GetUserInfo(string userId) {
            
            StringBuilder sSQL = new StringBuilder();
            SqlCommand command = null;
            SqlDataAdapter adapter = null;
            DataTable dt = null;

            UserInfo userInfo = null;

            EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
            SqlConnectionStringBuilder sqlsb;

            sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["myConnString"].ConnectionString);

            sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

            using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
            {
                try
                {
                    dt = new DataTable();                    

                    sSQL.Clear();
                    sSQL.AppendFormat("SELECT MV001, MV002, ME001, ME002 FROM dbo.CMSMV Join dbo.CMSME on MV004 = ME001 WHERE MV001 = '{0}'", userId);
                    conn.Open();

                    command = new SqlCommand(sSQL.ToString(), conn);
                    adapter = new SqlDataAdapter(command);
                    adapter.Fill(dt);

                    for (int i = 0; i < dt.Rows.Count; i++) {
                        userInfo = new UserInfo
                        {
                            UserID = dt.Rows[i]["MV001"].ToString().Trim(),
                            UserName = dt.Rows[i]["MV002"].ToString().Trim(),
                            DeptNo = dt.Rows[i]["ME001"].ToString().Trim(),
                            DeptName = dt.Rows[i]["ME002"].ToString().Trim()
                        };
                    }
                }
                catch (Exception e)
                {
                    _log.Debug("讀取使用者資料檔發生錯誤：" + e.Message);
                }
                finally
                {
                    if (adapter != null) adapter.Dispose();
                    if (command != null) command.Dispose();
                    if (dt != null) dt.Dispose();
                }
            }
            return userInfo;
        }

        /// <summary>
        /// 取得上一層的部門資訊
        /// </summary>
        /// <param name="userInfo"></param>
        /// <returns></returns>
        public static UserInfo GetParentDeptInfo(UserInfo userInfo, int deptLevel)
        {
            
            StringBuilder sSQL = new StringBuilder();
            SqlCommand command = null;
            SqlDataAdapter adapter = null;
            DataTable dt = null;
            UserInfo Parent = null;

            EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
            SqlConnectionStringBuilder sqlsb;

            sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["myConnString"].ConnectionString);

            sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

            using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
            {
                try
                {
                    dt = new DataTable();

                    sSQL.Clear();
                    
                    if (deptLevel == 1)
                        sSQL.AppendFormat("SELECT ME001, ME002 FROM dbo.CMSME WHERE ME001 like '{0}%' AND SUBSTRING(ME001, 3, 1) = 0 AND ME001 <> '600'", userInfo.DeptNo.Substring(0, 2));
                    else if (deptLevel == 2)
                        sSQL.AppendFormat("SELECT ME001, ME002 FROM dbo.CMSME WHERE ME001 like '{0}0%' AND SUBSTRING(ME001, 3, 1) = 0 AND ME001 <> '600'", userInfo.DeptNo.Substring(0, 1));
                    
                    conn.Open();

                    command = new SqlCommand(sSQL.ToString(), conn);
                    adapter = new SqlDataAdapter(command);
                    adapter.Fill(dt);

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        Parent = new UserInfo();
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            if (userInfo.DeptNo.Equals(dt.Rows[i]["ME001"].ToString().Trim()))
                                Parent = null;
                            else { 
                                Parent.DeptNo = dt.Rows[i]["ME001"].ToString().Trim();
                                Parent.DeptName = dt.Rows[i]["ME002"].ToString().Trim();
                            }
                        }
                    }                    
                }
                catch (Exception e)
                {
                    _log.Debug("讀取部門資料檔發生錯誤：" + e.Message);
                }
                finally
                {
                    if (adapter != null) adapter.Dispose();
                    if (command != null) command.Dispose();
                    if (dt != null) dt.Dispose();
                }
            }
            return Parent;
        }


        /// <summary>
		/// 取得登入者使用程式的權限
		/// </summary>
		/// <param name="userID">登入者代號</param>
		/// <returns></returns>
		public static string[] Get_DB_Auth(string userID)
        {
            string[] ANB_UserInfo = new string[3];
            StringBuilder sSQL = new StringBuilder();
            SqlCommand command = null;
            SqlDataAdapter adapter = null;
            DataTable dt = null;

            EncDec_Lib.EncDec Cryp = new EncDec_Lib.EncDec();
            SqlConnectionStringBuilder sqlsb;

            sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["AB_ConnString"].ConnectionString);

            sqlsb.Password = Cryp.DESDecrypt(sqlsb.Password);   // 解密

            using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
            {
                try
                {
                    dt = new DataTable();

                    sSQL.Clear();
                    sSQL.AppendFormat("SELECT TO002, TO003, TO004 FROM dbo.ANBTO WHERE (TO002 = '{0}' or TO004 = '{1}') AND TO005 <> 'Y'", userID, userID);
                    conn.Open();

                    command = new SqlCommand(sSQL.ToString(), conn);
                    adapter = new SqlDataAdapter(command);
                    adapter.Fill(dt);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {                        
                        ANB_UserInfo[0] = dt.Rows[i]["TO002"].ToString();   // 員工工號
                        ANB_UserInfo[1] = dt.Rows[i]["TO003"].ToString();   // 姓名                        
                        ANB_UserInfo[2] = dt.Rows[i]["TO004"].ToString();   // AD帳號
                    }
                }
                 
                catch (Exception e)
                {
                    _log.Debug("讀取ANBTO資料發生錯誤：" + e.Message);
                }
                finally
                {
                    if (adapter != null) adapter.Dispose();
                    if (command != null) command.Dispose();
                    if (dt != null) dt.Dispose();
                }
            }
            //return authority;
            return ANB_UserInfo;
        }
    }
}
