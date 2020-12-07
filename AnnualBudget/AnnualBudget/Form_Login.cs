using AnnualBudget.Model;
using AnnualBudget.BOs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using log4net.Config;

namespace AnnualBudget
{
    public partial class Form_Login : Form
    {
        private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private Form_Main custForm;
        public Form_Main CustForm { get => custForm; set => custForm = value; }

        public Form_Login()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
        }

        public Form_Login(Form_Main form1) {
            this.custForm = form1;
        }

        private void Form_Login_Load(object sender, EventArgs e)
        {
            //ModifyItem("", "", "", "");
        }

        private void btn_Login_Click(object sender, EventArgs e)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            _log.Debug("使用者登入中...");
            btn_Login.Text = "登入中...";
            btn_Login.Enabled = false;

            Form_Main form1;
            if (custForm != null)
                form1 = this.custForm;
            else
                form1 = new Form_Main();


            bool auth = false;
            UserInfo userInfo = new UserInfo();

            if (String.IsNullOrEmpty(tbx_userName.Text) || String.IsNullOrEmpty(tbx_PWD.Text))
            {
                _log.Debug("帳號或密碼欄位為空值");
                btn_Login.Enabled = true;
                btn_Login.Text = "登 入";
                MessageBox.Show("帳號或密碼欄不能為空值！");
            }
            else
            {
                auth = UserLogin(tbx_userName.Text, tbx_PWD.Text);  // 進行驗證

                if (auth == false)
                {

                    form1.gUserInfo.DeptNo = "";
                    btn_Login.Enabled = true;
                    btn_Login.Text = "登 入";
                    _log.Debug("AD帳號認證失敗：" + tbx_userName.Text);
                    MessageBox.Show("AD帳號認證失敗，請重新登入！");
                    return;

                }
                else
                {
                    string userId = "";
                    string userName = "";
                    string userAD_ID = "";
                    
                    string[] DB_UserInfo = Auth_Model.Get_DB_Auth(tbx_userName.Text);

                    
                    userId = DB_UserInfo[0];
                    userName = DB_UserInfo[1];
                    userAD_ID = DB_UserInfo[2];



                    if (userId == null || String.IsNullOrEmpty(userId))
                    {                        
                        //_log.Debug("ANBTO表無該使用者，使用AD權限登入：" + tbx_userName.Text);
                        //userInfo = Auth_Model.GetUserInfo(tbx_userName.Text);  // 抓取部門代號及名稱 

                        _log.Debug("沒有使用權限：" + tbx_userName.Text);

                        form1.gUserInfo.DeptNo = "";
                        btn_Login.Enabled = true;
                        btn_Login.Text = "登 入";

                        MessageBox.Show("很抱歉，您沒有使用權限，請重新確認後再登入");
                        return;

                    }
                    else {

                        //userId = "20150509";
                        _log.Debug("使用ANBTO的資料去ERP找部門資料：" + userId);
                        userInfo = Auth_Model.GetUserInfo(userId);  // 抓取部門代號及名稱                         

                        if (userInfo != null)
                        {                            
                            // 若登入者本身就屬於[部]級別的, ex: 830
                            if (Convert.ToInt32(userInfo.DeptNo.Substring(1, 2).ToString()) != 0 && Convert.ToInt32(userInfo.DeptNo.Last().ToString()) == 0)
                            {
                                userInfo.ParentDept = Auth_Model.GetParentDeptInfo(userInfo, 2);     // 抓取上一層[處]級別的部門代號及名稱
                                userInfo.Level = "2";
                            }
                            else if (Convert.ToInt32(userInfo.DeptNo.Last().ToString()) > 0)         // 若登入者本身就屬於[科]級別的, ex: 831  
                            {
                                userInfo.Level = "3";
                                userInfo.ParentDept = Auth_Model.GetParentDeptInfo(userInfo, 1);     // 抓取上一層[部]級別的部門代號及名稱
                                if (userInfo.ParentDept != null)
                                    userInfo.ParentDept.ParentDept = Auth_Model.GetParentDeptInfo(userInfo.ParentDept, 2);     // 再抓取上一層[處]級別的部門代號及名稱
                            }
                            else if(Convert.ToInt32(userInfo.DeptNo.Substring(1, 2).ToString()) == 0)   // 若是「處」級別的, ex: 800
                                userInfo.Level = "1";



                            form1.gUserInfo = userInfo;
                            //form1.gMainDeptNo = userInfo.DeptNo;
                            /*
                            ModifyItem(tbx_userName.Text, userInfo.UserName, userInfo.DeptNo, userInfo.DeptName);        // 記錄userID和部門代號
                            form1.gUserId = userInfo.UserID;
                            form1.gUserName = userInfo.UserName;
                            form1.gDeptNo = userInfo.DeptNo;
                            form1.gDeptName = userInfo.DeptName;

                            if (userInfo.ParentDept != null) { 
                                form1.gParentDeptNo = userInfo.ParentDept.DeptNo;
                                form1.gParentDeptName = userInfo.ParentDept.DeptName;
                            }
                            */
                            _log.Debug("登入成功，登入帳號：" + tbx_userName.Text + "，部門：" + userInfo.DeptNo);

                            this.DialogResult = DialogResult.Yes;
                        }
                        else
                        {
                            _log.Debug("沒有使用權限：" + tbx_userName.Text);

                            form1.gUserInfo.DeptNo = "";
                            btn_Login.Enabled = true;
                            btn_Login.Text = "登 入";

                            MessageBox.Show("很抱歉，您沒有使用權限，請重新確認後再登入");
                            return;
                        }
                    }
                }
            }
        }

        private bool UserLogin(string userID, string pwd) {
            bool auth = false;
            Auth_Model am = new Auth_Model();

            auth = am.ValidateKrtcAD(userID, pwd);
            return auth;
            
        }

        private void ModifyItem(string UserId, string UserName, string DeptNo, string DeptName)
        {

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            config.AppSettings.Settings["UserId"].Value = UserId;
            config.AppSettings.Settings["DeptNo"].Value = DeptNo;
            
            config.Save(ConfigurationSaveMode.Modified);

            ConfigurationManager.RefreshSection("appSettings");
        }

        private void Form_Login_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (this.DialogResult != DialogResult.Yes) 
                this.DialogResult = DialogResult.Cancel;
        }
    }
}
