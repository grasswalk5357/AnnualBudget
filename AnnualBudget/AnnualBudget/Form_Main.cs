using AnnualBudget.BOs;
using AnnualBudget.Model;
using log4net.Config;
using NPOI.SS.Formula.Functions;
using NPOI.XWPF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Configuration;

namespace AnnualBudget
{
    public partial class Form_Main : Form
    {
        private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

        Dictionary<string, decimal[]> allMonthlyData = new Dictionary<string, decimal[]>();

        public string[] RD_Depts = new string[] { "610", "900" };   // 開發部(610)、新創處(900)
        public UserInfo gUserInfo = new UserInfo();
        public List<Object> gList_depts = new List<Object>();
        public List<Object> gMonthlySumList = new List<object>();
        public string gToday = DateTime.Now.ToString("yyyyMMdd");
        public Stack<string> Ref_Id_Stack = new Stack<string>();
        /*
        public string gParentDeptNo = "";      // 父階部門代號
        public string gParentDeptName = "";    // 父階部門名稱
        public string gDeptNo = "";         // 部門代號
        public string gDeptName = "";       // 部門名稱
        public string gUserId = "";         // 登入者代號
        public string gUserName = "";       // 登入者姓名
        */
        int gDirectSalaryIndex = 0;          // 預算總表裡面的[直接人員薪資]的RowIndex
        int gDirectSalaryAddNumIndex = 0;    // 預算總表裡面的[直接人員薪資新增數]的RowIndex
        int gIndirectSalaryIndex = 0;        // 預算總表裡面的[間接人員薪資]的RowIndex
        int gIndirectSalaryAddNumIndex = 0;  // 預算總表裡面的[間接人員薪資新增數]的RowIndex
        int gPRJ_Index = -1;                 // 研發專案彙總表_目前所選取的RowIndex

        //public string gTempSelectedPrjIndex = "";
        public string gMainDeptNo = "";      // 記錄目前是在填寫哪個單位(代號)的預算表
        public string gMainDeptName = "";    // 記錄目前是在填寫哪個單位(名稱)的預算表

        public decimal gHealthIsr = 0;       // 健保費率
        public decimal gLaborIsr = 0;        // 勞保費率
        public decimal gFood = 0;            // 伙食費率
        public decimal gPension = 0;         // 退休金提撥率

        public List<object> summaryList = new List<object>();
        public static DataTable gDT_Summary = null;    // 記錄預算總表的DataTable

        bool isANBTC_Saved = false;  // [人力需求預算表]的儲存狀態
        bool isANBTD_Saved = false;  // [教育訓練計劃表]的儲存狀態
        bool isANBTE_Saved = false;  // [出差計劃表]的儲存狀態
        bool isANBTF_Saved = false;  // [資本支出預算表]的儲存狀態
        bool isANBTG_Saved = false;  // [研發專案彙總表]的儲存狀態
        bool isANBTI_Saved = false;  // [預算總表]的儲存狀態
        bool isANBTN_Saved = false;  // [組織編制表]的儲存狀態
        //private Object oDocument;


        Dictionary<string, Object> keyValues = new Dictionary<string, Object>();

        string gTmplID = "";         // 記錄樣版代號


        public Form_Main()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
        }

        private void Form_Main_Load(object sender, EventArgs e)
        {
            _log.Debug("主程式載入中...");

            Form_Login loginForm = new Form_Login();
            loginForm.CustForm = this;


            if (OpenForm(loginForm) == DialogResult.Yes)    // 開啟登入頁面，且資訊正確
            {
                Init_Data();
                /*
                lbl_PRJ_Notice.Text = "資料全部暫存後，" + Environment.NewLine + "記得按下方[儲存本表]，" + Environment.NewLine + "才能正確儲存至資料庫";

                SetYear();  // 設定年份
                gList_depts = ANBTK_Model.LoadDepts(gUserInfo, cbx_Year.Text, gList_depts);


                try {
                    if (gList_depts != null && gList_depts.Count > 0) {
                        for (int i = 0; i < gList_depts.Count; i++)
                            cbx_Dept.Items.Add(gList_depts[i]);

                        if (cbx_Dept.Items.Count > 1) 
                            btn_SwitchDept.Enabled = true;
                        else
                            btn_SwitchDept.Enabled = false;


                        //gMainDeptNo = cbx_Dept.Text;
                        gMainDeptNo = cbx_Dept.Items[0].ToString();
                        gDT_Summary = Tmpl_Model.GetTmplContentByDept(gMainDeptNo);
                        
                        GetTmplID();    // 取得全域的樣版編號

                        //dgv_Summary = Set_dgvSummary(gMainDeptNo, dgv_Summary, gDT_Summary);
                        dgv_Summary = ANBTI_Model.Set_dgvSummary(gDT_Summary, dgv_Summary, false, gMainDeptNo);
                    }
                }
                catch (Exception ex) { _log.Debug(ex.Message); }

                if (cbx_Dept.Items.Count > 0)
                    cbx_Dept.SelectedIndex = 0;

                _log.Debug("填寫部門代號為：" + gMainDeptNo + "的年度預算表");

                if (gMainDeptNo.Equals(gUserInfo.DeptNo)) 
                    gMainDeptName = gUserInfo.DeptName;
                else if (gUserInfo.ParentDept != null && gMainDeptNo.Equals(gUserInfo.ParentDept.DeptNo)) 
                    gMainDeptName = gUserInfo.ParentDept.DeptName;

                if (String.IsNullOrEmpty(gMainDeptNo)) {
                    MessageBox.Show("您不需填寫年度預算表，或沒有填寫的權限。\n" + "請按確定關閉程式");
                    this.Close();
                }
                else
                {
                    lockHumanType();
                    //Load_and_Set_DGV_Data(string annualBudgetFormID, DataGridView dgv, string DeptNo, string year, string TmplID)
                    //List<Object> list = Common_Model.LoadMatrixData(annualBudgetFormID, gMainDeptNo, cbx_Year.Text, gTmplID);                    

                    Load_and_Set_DGV_Data("ABM001", dgv_HumanBudget, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
                    Load_and_Set_DGV_Data("ABM002", dgv_EmployeeTraining, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
                    Load_and_Set_DGV_Data("ABM003", dgv_BusinessTrip, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
                    Load_and_Set_DGV_Data("ABM004", dgv_CapEx, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
                    LoadAndSet_ANBTG_DGV("ABM005", dgv_PRJ, dgv_PRJ_Content);
                    Load_and_Set_DGV_Data("ABM007", dgv_Summary, gMainDeptNo, cbx_Year.Text, gTmplID, gDT_Summary, false);
                    Load_and_Set_DGV_Data("ABM008", dgv_Org, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);  

                    Set_dgvPRJ_Content();

                    CreateMap();    // 建立總表內，自動帶出值的會科識別號
                    GetRate();      // 取得某些會科所需的費率
                    SetTabControlAuth();    

                    //tbx_MainDeptNo.Text = gMainDeptNo;

                    if (gUserInfo.ParentDept != null) { 
                        tbx_ParentDeptNo.Text = gUserInfo.ParentDept.DeptNo;
                        tbx_ParentDeptName.Text = gUserInfo.ParentDept.DeptName;
                    }

                    tbx_DeptNo.Text = gUserInfo.DeptNo;
                    tbx_DeptName.Text = gUserInfo.DeptName;

                    setTableNum();                   
                                        
                    OpenDeptButton();
                }
                */
            }
            else
                this.Close();

            _log.Debug("主程式載入完畢！");            
        }

        public DialogResult OpenForm(Form form)
        {
            form.ShowDialog();

            return form.DialogResult;
        }
        private void dgv_Org_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            SetRowsIndex(dgv_Org);
        }
        private void dgv_HumanBudget_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            SetRowsIndex(dgv_HumanBudget);
            /*DataGridViewComboBoxCell cbc = (DataGridViewComboBoxCell)dgv_HumanBudget.Rows[dgv_HumanBudget.RowCount - 1].Cells["dgv_HBT_cln_StartMonth"];
            if (cbc.Value == null || cbc.Items.Count <= 0) {
                for (int i = 0; i <= 12; i++)
                {                
                   cbc.Items.Add(i.ToString());
                }
            }
            else {
                for (int i = 0; i <= 12; i++) {
                    if (!cbc.Items.Contains(i))
                        cbc.Items.Add(i);
                }
            }*/
        }

        private void dgv_EmployeeTraining_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            SetRowsIndex(dgv_EmployeeTraining);
        }

        private void dgv_BusinessTrip_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            SetRowsIndex(dgv_BusinessTrip);
        }

        private void dgv_CapEx_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            SetRowsIndex(dgv_CapEx);            
        }

        public void SetRowsIndex(DataGridView dgv) {
            for (int i = 0; i < dgv.RowCount; i++)
                dgv.Rows[i].HeaderCell.Value = (i + 1).ToString();

            dgv.RowHeadersWidth = 60;
        }

        public void SetYear() {
            int year = DateTime.Today.Year;

            for (int i = year; i < year + 3; i++) {
                cbx_Year.Items.Add(i);
            }

            cbx_Year.SelectedIndex = 1;
        }

        private void dgv_Org_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int ActNum = 0;     // 實際人數
            int PlanNum = 0;     // 計畫人數
            int DiffNum = 0;    // 實際人數與計畫人數差額

            try {
                //if (dgv_Org.Columns["dgv_Org_cln_ActualNum"].Index == e.ColumnIndex)
                    ActNum = Convert.ToInt32(dgv_Org.Rows[e.RowIndex].Cells["dgv_Org_cln_ActNum"].Value);
                //if (dgv_Org.Columns["dgv_Org_cln_PlanNum"].Index == e.ColumnIndex)
                    PlanNum = Convert.ToInt32(dgv_Org.Rows[e.RowIndex].Cells["dgv_Org_cln_EstNum"].Value);

                DiffNum = PlanNum - ActNum;
                dgv_Org.Rows[e.RowIndex].Cells["dgv_Org_cln_DiffNum"].Value = DiffNum;
            }
            catch (Exception ex) {
                ExceptionHandle(dgv_Org, e.RowIndex, e.ColumnIndex, "公式計算錯誤：" + ex.Message, "部門編制表");
            }

            try
            {
                string ref_id = "";
                if (dgv_Org.Rows[e.RowIndex].Cells["dgv_Org_cln_Ref_ANBTN"].Value == null ||
                    String.IsNullOrEmpty(dgv_Org.Rows[e.RowIndex].Cells["dgv_Org_cln_Ref_ANBTN"].Value.ToString()))
                {
                    if (Ref_Id_Stack.Count <= 0)
                    {
                        ref_id = ANBTN_Model.Get_Ref_TNID();                        
                        Ref_Id_Stack.Push(ref_id);
                    }
                    else
                    {
                        ref_id = Ref_Id_Stack.Peek();
                        ref_id = (Convert.ToInt32(ref_id) + 1).ToString();
                        Ref_Id_Stack.Push(ref_id);                        
                    }
                    dgv_Org.Rows[e.RowIndex].Cells["dgv_Org_cln_Ref_ANBTN"].Value = ref_id;

                }
            }
            catch (Exception ex) {
                ExceptionHandle(dgv_Org, e.RowIndex, e.ColumnIndex, "產生參照編號錯誤：" + ex.Message , "部門編制表");
            }
            
        }

        private void dgv_HumanBudget_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            decimal ChangeNum = 0;  // 增減人數
            decimal ActNum = 0;     // 實際人數
            decimal EstNum = 0;     // 計畫人數
            decimal Salary = 0;     // 增減之人員每月薪資
            decimal TotalChangeSalary = 0;    // 每月增減薪資 
            int[] numIndex = new int[] { 4, 5, 8 };   // 記錄須填入數字的欄位位置

            try
            {
                if (numIndex.Contains(e.ColumnIndex))
                {
                    dgv_HumanBudget.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = ConvertToDecimal(dgv_HumanBudget.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);

                    ActNum = Convert.ToDecimal(dgv_HumanBudget.Rows[e.RowIndex].Cells["dgv_HBT_cln_ActNum"].Value);    // 實際人數
                    EstNum = Convert.ToDecimal(dgv_HumanBudget.Rows[e.RowIndex].Cells["dgv_HBT_cln_EstNum"].Value);    // 計畫人數

                    ChangeNum = EstNum - ActNum;    // 增減人數 = 計畫人數 - 實際人數

                    dgv_HumanBudget.Rows[e.RowIndex].Cells["dgv_HBT_cln_DiffNum"].Value = ChangeNum;

                    Salary = Convert.ToDecimal(dgv_HumanBudget.Rows[e.RowIndex].Cells["dgv_HBT_cln_Salary"].Value);
                    TotalChangeSalary = Math.Round(Salary * ChangeNum, 0, MidpointRounding.AwayFromZero);  // 每月增減薪資 = 增減人數 * 增減之人員每月薪資
                    dgv_HumanBudget.Rows[e.RowIndex].Cells["dgv_HBT_cln_TotalChangeSalary"].Value = TotalChangeSalary;
                }

                ANBTC_Model.CheckDataValue(dgv_HumanBudget);

            }            
            catch(Exception ex) 
            {
                ExceptionHandle(dgv_HumanBudget, e.RowIndex, e.ColumnIndex, ex.Message, "人力需求預算表");
            }
        }
        private void dgv_EmployeeTraining_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            decimal StartMonth = 0;   // 起始月
            decimal EndMonth = 0;     // 結束月
            decimal MonthlyFee = 0;   // 每月費用
            decimal ProjectQuota = 0; // 該計劃費用總額
            int[] numIndex = new int[] { 3, 7, 8 };   // 記錄須填入數字的欄位位置
            try
            {
                if (numIndex.Contains(e.ColumnIndex))
                {
                    dgv_EmployeeTraining.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = ConvertToDecimal(dgv_EmployeeTraining.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);

                    StartMonth = Convert.ToDecimal(dgv_EmployeeTraining.Rows[e.RowIndex].Cells[5].Value);
                    EndMonth = Convert.ToDecimal(dgv_EmployeeTraining.Rows[e.RowIndex].Cells[6].Value);

                    if (EndMonth < StartMonth)
                    {
                        MessageBox.Show("結束月不可以小於起始月！");
                        return;
                    }

                    MonthlyFee = Convert.ToDecimal(dgv_EmployeeTraining.Rows[e.RowIndex].Cells[8].Value);

                    ProjectQuota = Math.Round((EndMonth - StartMonth + 1) * MonthlyFee, 0, MidpointRounding.AwayFromZero);     // 該計劃費用總額 = (結束月 - 起始月 + 1) * MonthlyFee            

                    dgv_EmployeeTraining.Rows[e.RowIndex].Cells[9].Value = ProjectQuota;
                }
            }
            catch (Exception ex) {
                ExceptionHandle(dgv_EmployeeTraining, e.RowIndex, e.ColumnIndex, ex.Message, "教育訓練計劃表");
            }
        }

        private void dgv_BusinessTrip_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            decimal airFare = 0;            // 機票款
            decimal hotelFare = 0;          // 住宿費
            decimal shippingExpenses = 0;   // 交通費
            decimal otherFare = 0;          // 雜費
            decimal dailyExpense = 0;       // 日支費
            decimal foodStipend = 0;        // 餐費
            decimal sumOfTripExpense = 0;   // 旅費小計
            int[] numIndex = new int[] { 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };   // 記錄須填入數字的欄位位置
            try
            {
                if (numIndex.Contains(e.ColumnIndex))
                {
                    dgv_BusinessTrip.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = ConvertToDecimal(dgv_BusinessTrip.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);


                    airFare = Convert.ToDecimal(dgv_BusinessTrip.Rows[e.RowIndex].Cells[4].Value);
                    hotelFare = Convert.ToDecimal(dgv_BusinessTrip.Rows[e.RowIndex].Cells[5].Value);
                    shippingExpenses = Convert.ToDecimal(dgv_BusinessTrip.Rows[e.RowIndex].Cells[6].Value);
                    otherFare = Convert.ToDecimal(dgv_BusinessTrip.Rows[e.RowIndex].Cells[7].Value);
                    dailyExpense = Convert.ToDecimal(dgv_BusinessTrip.Rows[e.RowIndex].Cells[8].Value);
                    foodStipend = Convert.ToDecimal(dgv_BusinessTrip.Rows[e.RowIndex].Cells[9].Value);

                    sumOfTripExpense = airFare + hotelFare + shippingExpenses + otherFare + dailyExpense + foodStipend;
                    dgv_BusinessTrip.Rows[e.RowIndex].Cells[10].Value = sumOfTripExpense;
                }
            }
            catch (Exception ex) {
                ExceptionHandle(dgv_BusinessTrip, e.RowIndex, e.ColumnIndex, ex.Message, "出差計劃表");
            }
        }

        private void dgv_CapEx_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            decimal num = 0;
            decimal unitPrice = 0;
            decimal totalPrice = 0;
            decimal assetLife = 0;
            decimal depre = 0;
            int month = 0;

            int[] numIndex = new int[] { 3, 4, 6, 7, 9 };   // 記錄須填入數字的欄位位置
            try
            {
                if (numIndex.Contains(e.ColumnIndex))
                {
                    dgv_CapEx.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = ConvertToDecimal(dgv_CapEx.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);


                    num = Convert.ToDecimal(dgv_CapEx.Rows[e.RowIndex].Cells[3].Value);         // 數量
                    unitPrice = Convert.ToDecimal(dgv_CapEx.Rows[e.RowIndex].Cells[4].Value);   // 單價

                    if (num > 0 && unitPrice > 0)
                        totalPrice = num * unitPrice;   // 總價金額

                    dgv_CapEx.Rows[e.RowIndex].Cells[5].Value = totalPrice;

                    month = Convert.ToInt32(dgv_CapEx.Rows[e.RowIndex].Cells[8].Value);

                    assetLife = Convert.ToDecimal(dgv_CapEx.Rows[e.RowIndex].Cells[6].Value);   // 耐用年限

                    if (totalPrice != 0 && assetLife != 0)
                        depre = Math.Round(totalPrice / assetLife / 12, 0, MidpointRounding.AwayFromZero);     // 每月折舊

                    dgv_CapEx.Rows[e.RowIndex].Cells[10].Value = depre;
                }
            }
            catch (Exception ex) {
                ExceptionHandle(dgv_CapEx, e.RowIndex, e.ColumnIndex, ex.Message, "資本支出預算表");
            }
        }

        private void btn_SaveOrg_Click(object sender, EventArgs e)
        {
            _log.Debug("儲存[組織編制表]中...");
            string annualBudgetFormID = "ABM008";

            ANBTB anbtb = new ANBTB(gUserInfo.DeptNo, gMainDeptNo, cbx_Year.Text, annualBudgetFormID, "組織編制表", gUserInfo.UserID, DateTime.Now.ToString("yyyyMMdd"));
                        
            List<Object> myList = ANBTN_Model.SaveMatrixData(dgv_Org, annualBudgetFormID, gMainDeptNo, gUserInfo.DeptNo, cbx_Year.Text, gUserInfo.UserID, gToday);

            ANBTN_Model.SyncDataToDGV_HBT(myList, dgv_HumanBudget);

            ANBTC_Model.SetSyncRowColor(dgv_HumanBudget);
            //List<decimal[]> dList = ANBTC_Model.GetMonthlyData(myList);

            //allMonthlyData["1001"] = dList[0];     // 人力需求預算表-增減薪資金額-直接人員
            //allMonthlyData["1002"] = dList[1];     // 人力需求預算表-增減薪資金額-間接人員

            Set_dgvSummaryContent(dgv_Summary, gDT_Summary, false);

            if (myList.Count >= 1)
                anbtb.Tb006 = "Y";
            else
                anbtb.Tb006 = "N";

            if (ANBTB_Model.UpdateToANBTB(anbtb) == false || ANBTN_Model.UpdateToANBTN(myList) == false)
            {
                _log.Debug("儲存[組織編制表]至DB時發生錯誤！");
                MessageBox.Show("儲存[組織編制表]至DB時發生錯誤！");
            }
            else
            {
                if (isANBTN_Saved == false)
                {
                    isANBTN_Saved = true;
                    this.PageA.Text = "✔" + this.PageA.Text; // 將標籤標題加上✔號，代表已儲存
                }

                Load_and_Set_DGV_Data("ABM008", dgv_Org, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);  // 將資料重新載入


                _log.Debug("儲存[組織編制表]成功！重新載入資料");

                this.tabControl1.SelectedTab = this.PageB;
                //Summarise(-1);            // [預算總表]計算各月份的總計
                //ComputeMonthlySum();    // [預算總表]計算各月份的總計
            }
        }

        private void btn_SaveHumanBudget_Click(object sender, EventArgs e)
        {
            _log.Debug("儲存[人力需求預算表]中...");
            string annualBudgetFormID = "ABM001";

            bool checkResult = ANBTC_Model.CheckMatrixData(dgv_HumanBudget);    // 檢核必填欄位是否已填

            if (checkResult == true)
            {
                ANBTB anbtb = new ANBTB(gUserInfo.DeptNo, gMainDeptNo, cbx_Year.Text, annualBudgetFormID, "人力需求預算表", gUserInfo.UserID, DateTime.Now.ToString("yyyyMMdd"));

                //List<Object> myList = HumanBudget_Model.SaveMatrixData(dgv_HumanBudget, annualBudgetFormID, gUserInfo.DeptNo, cbx_Year.Text);
                List<Object> myList = ANBTC_Model.SaveMatrixData(dgv_HumanBudget, annualBudgetFormID, gMainDeptNo, gUserInfo.DeptNo, cbx_Year.Text, gUserInfo.UserID, gToday);

                List<decimal[]> dList = ANBTC_Model.GetMonthlyData(myList);

                allMonthlyData["1001"] = dList[0];     // 人力需求預算表-增減薪資金額-直接人員
                allMonthlyData["1002"] = dList[1];     // 人力需求預算表-增減薪資金額-間接人員
                allMonthlyData["8002"] = dList[2];     // 伙食人數-人數小計(直接人員)
                allMonthlyData["8001"] = dList[3];     // 伙食人數-人數小計(間接人員)

                Set_dgvSummaryContent(dgv_Summary, gDT_Summary, false);

                if (myList.Count >= 1)
                    anbtb.Tb006 = "Y";
                else
                    anbtb.Tb006 = "N";

                if (ANBTB_Model.UpdateToANBTB(anbtb) == false || ANBTC_Model.UpdateToANBTC(myList) == false)
                {
                    _log.Debug("儲存[人力需求預算表]至DB時發生錯誤！");
                    MessageBox.Show("儲存[人力需求預算表]至DB時發生錯誤！");
                }
                else
                {
                    if (isANBTC_Saved == false)
                    {
                        isANBTC_Saved = true;
                        this.PageB.Text = "✔" + this.PageB.Text; // 將標籤標題加上✔號，代表已儲存
                    }

                    Load_and_Set_DGV_Data("ABM001", dgv_HumanBudget, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);  // 將資料重新載入
                    
                    ANBTC_Model.SetSyncRowColor(dgv_HumanBudget);

                    _log.Debug("儲存[人力需求預算表]成功！重新載入資料");

                    this.tabControl1.SelectedTab = this.PageC;
                    Summarise(-1);            // [預算總表]計算各月份的總計
                                              //ComputeMonthlySum();    // [預算總表]計算各月份的總計
                }
            }
            else {
                MessageBox.Show("有資料尚未完善，請再次檢查！\n請將有'**'的欄位完善填寫");
            }
        }

        private void btn_SaveEmployeeTraining_Click(object sender, EventArgs e)
        {
            _log.Debug("儲存[教育訓練計畫表]中...");
            string annualBudgetFormID = "ABM002";

            ANBTB anbtb = new ANBTB(gUserInfo.DeptNo, gMainDeptNo, cbx_Year.Text, annualBudgetFormID, "教育訓練計畫表", gUserInfo.UserID, DateTime.Now.ToString("yyyyMMdd"));

            List<Object> myList = ANBTD_Model.SaveMatrixData(dgv_EmployeeTraining, annualBudgetFormID, gMainDeptNo, gUserInfo.DeptNo, cbx_Year.Text, gUserInfo.UserID, gToday);

            allMonthlyData["2001"] = ANBTD_Model.GetMonthlyData(myList);     // 教育訓練計劃表-訓練費
            Set_dgvSummaryContent(dgv_Summary, gDT_Summary, false);


            if (myList.Count >= 1)
                anbtb.Tb006 = "Y";
            else
                anbtb.Tb006 = "N";

            if (ANBTB_Model.UpdateToANBTB(anbtb) == false || ANBTD_Model.UpdateToANBTD(myList) == false)
            {
                _log.Debug("儲存[教育訓練計畫表]至DB時發生錯誤！");
                MessageBox.Show("儲存[教育訓練計畫表]至DB時發生錯誤！");
            }
            else
            {
                if (isANBTD_Saved == false)
                {
                    isANBTD_Saved = true;
                    this.PageC.Text = "✔" + this.PageC.Text; // 將標籤標題加上✔號，代表已儲存
                }

                _log.Debug("儲存[教育訓練計畫表]成功！重新載入資料");
                Load_and_Set_DGV_Data("ABM002", dgv_EmployeeTraining, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);  // 將資料重新載入
                
                this.tabControl1.SelectedTab = this.PageD;
                Summarise(-1);            // [預算總表]計算各月份的總計
                //ComputeMonthlySum();    // [預算總表]計算各月份的總計
            }
        }

        private void btn_SaveBusinessTrip_Click(object sender, EventArgs e)
        {
            _log.Debug("儲存[出差計劃表]中...");
            string annualBudgetFormID = "ABM003";
            ANBTB anbtb = new ANBTB(gUserInfo.DeptNo, gMainDeptNo, cbx_Year.Text, annualBudgetFormID, "出差計劃表", gUserInfo.UserID, DateTime.Now.ToString("yyyyMMdd"));
            
            List<Object> myList = ANBTE_Model.SaveMatrixData(dgv_BusinessTrip, annualBudgetFormID, gMainDeptNo, gUserInfo.DeptNo, cbx_Year.Text, gUserInfo.UserID, gToday);
            List<decimal[]> dList = ANBTE_Model.GetMonthlyData(myList);


            allMonthlyData["3001"] = dList[0];      // 出差計劃表-差旅費
            allMonthlyData["3002"] = dList[1];      // 出差計劃表-交際費
            allMonthlyData["3003"] = dList[2];      // 出差計劃表-旅平險
            Set_dgvSummaryContent(dgv_Summary, gDT_Summary, false);     //


            if (myList.Count >= 1)
                anbtb.Tb006 = "Y";
            else
                anbtb.Tb006 = "N";

            if (ANBTB_Model.UpdateToANBTB(anbtb) == false || ANBTE_Model.UpdateToANBTE(myList) == false)
            {
                _log.Debug("儲存[出差計劃表]至DB時發生錯誤！");
                MessageBox.Show("儲存[出差計劃表]至DB時發生錯誤！");
            }
            else
            {
                if (isANBTE_Saved == false)
                {
                    isANBTE_Saved = true;
                    this.PageD.Text = "✔" + this.PageD.Text; // 將標籤標題加上✔號，代表已儲存
                }
                _log.Debug("儲存[出差計劃表]成功！重新載入資料");
                Load_and_Set_DGV_Data("ABM003", dgv_BusinessTrip, gMainDeptNo, cbx_Year.Text, gTmplID, null, false); // 將資料重新載入
                
                this.tabControl1.SelectedTab = this.PageE;
                Summarise(-1);            // [預算總表]計算各月份的總計
                //ComputeMonthlySum();    // [預算總表]計算各月份的總計
            }
        }

        private void btn_SaveCapEx_Click(object sender, EventArgs e)
        {
            _log.Debug("儲存[資本支出預算表]中...");
            string annualBudgetFormID = "ABM004";
            ANBTB anbtb = new ANBTB(gUserInfo.DeptNo, gMainDeptNo, cbx_Year.Text, annualBudgetFormID, "資本支出預算表", gUserInfo.UserID, DateTime.Now.ToString("yyyyMMdd"));

            List<Object> myList = ANBTF_Model.SaveMatrixData(dgv_CapEx, annualBudgetFormID, gMainDeptNo, gUserInfo.DeptNo, cbx_Year.Text, gUserInfo.UserID, gToday);
            List<decimal[]> dList = ANBTF_Model.GetMonthlyData(myList, dgv_CapEx);

            allMonthlyData["4001"] = dList[0];      // 資本支出預算表-折舊
            allMonthlyData["4002"] = dList[1];      // 資本支出預算表-攤銷 
            Set_dgvSummaryContent(dgv_Summary, gDT_Summary, false);


            if (myList.Count >= 1)
                anbtb.Tb006 = "Y";
            else
                anbtb.Tb006 = "N";

            if (ANBTB_Model.UpdateToANBTB(anbtb) == false || ANBTF_Model.UpdateToANBTF(myList) == false)
            {
                _log.Debug("儲存[資本支出預算表]至DB時發生錯誤！");
                MessageBox.Show("儲存[資本支出預算表]至DB時發生錯誤！");
            }
            else
            {
                if (isANBTF_Saved == false)
                {
                    isANBTF_Saved = true;
                    this.PageE.Text = "✔" + this.PageE.Text; // 將標籤標題加上✔號，代表已儲存                    
                }
                _log.Debug("儲存[資本支出預算表]成功！重新載入資料");
                Load_and_Set_DGV_Data("ABM004", dgv_CapEx, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);        // 將資料重新載入
                

                if (RD_Depts.Contains(gUserInfo.DeptNo))     // 610:開發部。900:新創處
                    this.tabControl1.SelectedTab = this.PageF;
                else
                    this.tabControl1.SelectedTab = this.PageG;
                
                Summarise(-1);            // [預算總表]計算各月份的總計
                //ComputeMonthlySum();    // [預算總表]計算各月份的總計
            }
        }

        private void btn_SavePRJ_Click(object sender, EventArgs e)
        {
            if (gPRJ_Index < 0) {
                MessageBox.Show("請先點選要進行編輯或儲存的專案，再進行下一步");
            }
            else { 

                int rowIndex = dgv_PRJ.SelectedCells[0].RowIndex;
                bool result1 = CheckIsDate(rowIndex, dgv_PRJ.Rows[rowIndex].Cells["cln_StartPRJ"].Value, "cln_StartPRJ");
                bool result2 = CheckIsDate(rowIndex, dgv_PRJ.Rows[rowIndex].Cells["cln_EndPRJ"].Value, "cln_EndPRJ");
                bool result3 = CheckIsDate(rowIndex, dgv_PRJ.Rows[rowIndex].Cells["cln_TR_Day"].Value, "cln_TR_Day");


                if (result1 == true && result2 == true && result3 == true)
                {
                    SavePRJ_Data(dgv_PRJ.SelectedCells[0].RowIndex);
                    lbl_PRJ_Notice.Text = "已編輯的專案序號 [" + dgv_PRJ.Rows[gPRJ_Index].Cells[0].Value + "] 儲存成功";
                }
                else {
                    MessageBox.Show("日期格式不正確，請重新檢查。日期格式必須為yyyyMMdd。例：20201027");
                }
            }
        }

        private void btn_AddPRJ_Click(object sender, EventArgs e)
        {
                       
            if (dgv_PRJ.RowCount > 0) 
            {
                if (dgv_PRJ.SelectedCells[0] != null)
                    SavePRJ_Data(dgv_PRJ.SelectedCells[0].RowIndex);
                if (dgv_PRJ_Content.Visible == false)
                    dgv_PRJ_Content.Visible = true;
            }
                

            ADD_DGV_PRJ_Row();  // 先幫單頭新增一筆row

            dgv_PRJ.ClearSelection();
            dgv_PRJ.Rows[dgv_PRJ.RowCount - 1].Cells[0].Selected = true;    // 新增一筆之後，將指標指向新的那一筆
            //gPRJ_Index = dgv_PRJ.SelectedCells[0].RowIndex;


            btn_SavePRJ.Enabled = true;

            SwithchPRJ_Index();

        }
        private void btn_SaveRD_PRJ_Click(object sender, EventArgs e)
        {

            _log.Debug("儲存[研發專案彙總表]中...");
            int rowIndex = dgv_PRJ.SelectedCells[0].RowIndex;
            bool result1 = CheckIsDate(rowIndex, dgv_PRJ.Rows[rowIndex].Cells["cln_StartPRJ"].Value, "cln_StartPRJ");
            bool result2 = CheckIsDate(rowIndex, dgv_PRJ.Rows[rowIndex].Cells["cln_EndPRJ"].Value, "cln_EndPRJ");
            bool result3 = CheckIsDate(rowIndex, dgv_PRJ.Rows[rowIndex].Cells["cln_TR_Day"].Value, "cln_TR_Day");


            if (result1 == true && result2 == true && result3 == true)
            {                

                SavePRJ_Data(dgv_PRJ.SelectedCells[0].RowIndex);    // 先儲存最後一筆停留的研發專案資料


                bool isAllDataSaved = true;

                gPRJ_Index = -1; // 設回預設值

                for (int i = 0; i < dgv_PRJ.Rows.Count; i++) 
                {
                    if (dgv_PRJ.Rows[i].HeaderCell.Value != null) {
                        isAllDataSaved = false;
                        break;
                    }            
                }

                if (isAllDataSaved == true)
                {

                    string annualBudgetFormID = "ABM005";

                    ANBTB anbtb = new ANBTB(gUserInfo.DeptNo, gMainDeptNo, cbx_Year.Text, annualBudgetFormID, "研發專案彙總表", gUserInfo.UserID, DateTime.Now.ToString("yyyyMMdd"));

                    //keyValues = ANBTG_Model.SaveMatrixData(keyValues, dgv_PRJ_Content, dgv_PRJ.Rows[dgv_PRJ.SelectedCells[0].RowIndex], annualBudgetFormID, gUserInfo.DeptNo, cbx_Year.Text);
                    List<decimal[]> dList = ANBTG_Model.GetMonthlyData(keyValues);
                    allMonthlyData["5001"] = dList[0];      // 研發專案彙總表-加班費 
                    allMonthlyData["5002"] = dList[1];      // 研發專案彙總表-樣品費
                    //allMonthlyData["4002"] = dList[2];    // 資本支出預算表-攤銷 
                    allMonthlyData["5003"] = dList[3];      // 研發專案彙總表-技術移轉、設計費(勞務費)
                    allMonthlyData["5007"] = dList[4];      // 研發專案彙總表-運費
                    allMonthlyData["5004"] = dList[5];      // 研發專案彙總表-認證費
                    allMonthlyData["5005"] = dList[6];      // 研發專案彙總表-模具(未滿8萬)及治具費
                    allMonthlyData["5006"] = dList[7];      // 研發專案彙總表-模具(8萬以上)
                    allMonthlyData["6001"] = dList[9];      // 資本支出預算表-攤銷
                    Set_dgvSummaryContent(dgv_Summary, gDT_Summary, false);


                    if (keyValues.Count >= 1)
                        anbtb.Tb006 = "Y";
                    else
                        anbtb.Tb006 = "N";

                    if (ANBTB_Model.UpdateToANBTB(anbtb) == false || ANBTG_Model.UpdateToANBTG(keyValues) == false)
                    {
                        _log.Debug("儲存[研發專案彙總表]至DB時發生錯誤！");
                        MessageBox.Show("儲存[研發專案彙總表]至DB時發生錯誤！");
                    }
                    else
                    {
                        if (isANBTG_Saved == false)
                        {
                            isANBTG_Saved = true;
                            this.PageF.Text = "✔" + this.PageF.Text; // 將標籤標題加上✔號，代表已儲存
                        }
                        _log.Debug("儲存[研發專案彙總表]成功！重新載入資料");
                        keyValues = RemoveDeletedObject(keyValues);                 // 將標記刪除的專案從Dictionary中移除
                        LoadAndSet_ANBTG_DGV("ABM005", dgv_PRJ, dgv_PRJ_Content);   // 將資料重新載入

                        lbl_PRJ_Notice.Text = "資料全部暫存後，" + Environment.NewLine + "記得按下方[儲存本表]，" + Environment.NewLine + "才能正確儲存至資料庫";

                        this.tabControl1.SelectedTab = this.PageG;               // 跳到下一頁
                        Summarise(-1);            // [預算總表]計算各月份的總計
                        //ComputeMonthlySum();    // [預算總表]計算各月份的總計
                    }
                }
                else {                
                    MessageBox.Show("本表(研發專案彙總表)尚有資料未暫存，請重新確認");                
                }
            }
            //else
            //{
            //    MessageBox.Show("日期格式不正確，請重新檢查。");
            //}
        }


        private void btn_Save_Summary_Click(object sender, EventArgs e)
        {
            _log.Debug("儲存[預算總表]中...");
            string annualBudgetFormID = "ABM007";

            ANBTB anbtb = new ANBTB(gUserInfo.DeptNo, gMainDeptNo, cbx_Year.Text, annualBudgetFormID, "費用預算表", gUserInfo.UserID, DateTime.Now.ToString("yyyyMMdd"));

            List<object> myList = ANBTI_Model.SaveMatrixData(dgv_Summary, gDT_Summary, annualBudgetFormID, gMainDeptNo, gUserInfo.DeptNo, cbx_Year.Text, gUserInfo.UserID, gToday);

            summaryList = myList;

            if (ANBTB_Model.UpdateToANBTB(anbtb) == false || ANBTI_Model.UpdateToANBTI(myList) == false)
            {
                _log.Debug("儲存「預算總表」至DB時發生錯誤！");
                MessageBox.Show("儲存「預算總表」至DB時發生錯誤！");
            }
            else
            {
                if (isANBTI_Saved == false)
                {
                    isANBTI_Saved = true;
                    this.PageG.Text = "✔" + this.PageG.Text; // 將標籤標題加上✔號，代表已儲存
                }
                _log.Debug("儲存[預算總表]成功！重新載入資料");
                Load_and_Set_DGV_Data("ABM007", dgv_Summary, gMainDeptNo, cbx_Year.Text, gTmplID, gDT_Summary, false);      // 將資料重新載入
                MessageBox.Show(cbx_Year.Text + "年度各預算表已填寫完畢，現可匯出Excel檔，或請直接關閉程式。");
            }
        }


        /// <summary>
        /// 設定預算總表樣式
        /// </summary>
        public DataGridView Set_dgvSummary(string deptNo, DataGridView dgv, DataTable table) {

            DataGridView tempDgv = dgv;

            //gDT_Summary = Tmpl_Model.GetTmplContentByDept(deptNo);

            // 設定每一個會科項目
            //for (int i = 0; i < gDT_Summary.Rows.Count; i++)
            for (int i = 0; i < table.Rows.Count; i++)
            {
                /*
                gTmplID = gDT_Summary.Rows[i]["TL002"].ToString();      // 樣版編號
                string c1 = gDT_Summary.Rows[i]["會科編號"].ToString();  // 會科編號
                string c2 = gDT_Summary.Rows[i]["會科名稱"].ToString();  // 會科名稱
                string c3 = gDT_Summary.Rows[i]["會科說明"].ToString();  // 會科名稱說明
                */
                gTmplID = table.Rows[i]["TL002"].ToString();      // 樣版編號
                string c1 = table.Rows[i]["會科編號"].ToString();  // 會科編號
                string c2 = table.Rows[i]["會科名稱"].ToString();  // 會科名稱
                string c3 = table.Rows[i]["會科說明"].ToString();  // 會科名稱說明
                /*dgv_Summary.Rows.Add();

                if (String.IsNullOrEmpty(c1))
                    dgv_Summary.Rows[i].HeaderCell.Value = c2 + "-(" + c3 + ")";
                else
                    dgv_Summary.Rows[i].HeaderCell.Value = c1 + "-" + c2 + "-(" + c3 + ")";
                */
                tempDgv.Rows.Add();

                if (String.IsNullOrEmpty(c1))
                    tempDgv.Rows[i].HeaderCell.Value = c2 + "-(" + c3 + ")";
                else
                    tempDgv.Rows[i].HeaderCell.Value = c1 + "-" + c2 + "-(" + c3 + ")";
            }

            //dgv_Summary.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;            
            tempDgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

            return tempDgv;
        }

        /// <summary>
        /// 設定研發專案彙總表-單身的DataGridView
        /// </summary>
        private void Set_dgvPRJ_Content() {

            // 設定每一個會科項目
            dgv_PRJ_Content.Rows.Clear();
            for (int i = 0; i < 9; i++)
                dgv_PRJ_Content.Rows.Add();

            dgv_PRJ_Content.Rows[0].HeaderCell.Value = "(1)加班費";
            dgv_PRJ_Content.Rows[1].HeaderCell.Value = "(2)樣品費(消耗器材及原料使用費)";
            dgv_PRJ_Content.Rows[2].HeaderCell.Value = "(3)樣品費(Mockup)";    
            dgv_PRJ_Content.Rows[3].HeaderCell.Value = "(4)技術移轉、設計費(勞務費)";
            dgv_PRJ_Content.Rows[4].HeaderCell.Value = "(5)運費";              
            dgv_PRJ_Content.Rows[5].HeaderCell.Value = "(6)認證費";            
            dgv_PRJ_Content.Rows[6].HeaderCell.Value = "(7)模具費(未滿8萬元)";
            dgv_PRJ_Content.Rows[7].HeaderCell.Value = "(8)模具(金額>=8萬元且耐用2年以上)";
            dgv_PRJ_Content.Rows[8].HeaderCell.Value = "(9)治具費(未滿8萬元)";

            //[(3)樣品費]和[(5)運費]介面上隱藏，暫時不使用
            //dgv_PRJ_Content.Rows[2].Visible = false;
            //dgv_PRJ_Content.Rows[4].Visible = false;

            dgv_PRJ_Content.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

        }

        /// <summary>
        /// 設定總表的DataGridView內容
        /// </summary>
        /// <param name="dgv"></param>
        private void Set_dgvSummaryContent(DataGridView dgv, DataTable TmplTable, bool isDeptAccounting)
        {
            string sheetCode = "";
            //string[] prj_item = new string[] {"5001", "5002", "5003", "5004", "5005", "5006", "5007", "6001" };

            //if (dgv != null && dgv.RowCount > 0)
            if (TmplTable != null && TmplTable.Rows.Count > 0)
            {
                for (int i = 0; i < TmplTable.Rows.Count; i++)
                {
                    string itemName = TmplTable.Rows[i]["會科說明"].ToString();
                    
                    // 先抓[直接人員][間接人員]的Index
                    if ("直接人員薪資".Equals(itemName))
                        gDirectSalaryIndex = i;
                    else if ("直接人員薪資新增數".Equals(itemName))
                        gDirectSalaryAddNumIndex = i;
                    else if ("間接人員薪資".Equals(itemName) || "薪資".Equals(itemName))
                        gIndirectSalaryIndex = i;
                    else if ("間接人員薪資新增數".Equals(itemName) || "薪資新增數".Equals(itemName))
                        gIndirectSalaryAddNumIndex = i;


                    //if (!String.IsNullOrEmpty(TmplTable.Rows[i]["TL010"].ToString()))     // TL010 是否需要手動輸入
                    if (isDeptAccounting == false) { 
                        if ("Y".Equals(TmplTable.Rows[i]["TL010"].ToString()))     // TL010 = Y：記錄是否為系統自動帶入值
                        {
                            dgv.Rows[i].DefaultCellStyle.BackColor = Color.PaleGreen;
                            dgv.Rows[i].ReadOnly = true;

                            sheetCode = TmplTable.Rows[i]["TL011"].ToString();


                            if (allMonthlyData.ContainsKey(sheetCode) == true && allMonthlyData[sheetCode] != null) // 將前幾頁分頁報表的每月統計數字帶進來
                            {
                                for (int j = 0; j < allMonthlyData[sheetCode].Length; j++)
                                {
                                    dgv.Rows[i].Cells[j].Value = allMonthlyData[sheetCode][j].ToString();
                                    dgv.Rows[i].Cells[j].Value = ConvertToDecimal(dgv.Rows[i].Cells[j].Value);
                                }
                            }
                        }
                        else if ("M".Equals(TmplTable.Rows[i]["TL010"].ToString())) // TL010 = M：記錄是否為財會部協助預先填寫的項目
                        {
                            dgv.Rows[i].DefaultCellStyle.BackColor = Color.PaleTurquoise;
                            dgv.Rows[i].ReadOnly = true;
                        }
                    }
                }
            }
        }


        /// <summary>
        /// 判斷輸入之參數是否為Decimal型態
        /// </summary>
        /// <param name="var"></param>
        /// <returns></returns>
        private bool isDecimal(string var) {
            bool result;

            if (decimal.TryParse(var, out decimal dVar)) {
                result = true;
            }
            else {
                MessageBox.Show("請輸入數字！");
                result = false;
            }
            return result;
        }


        private void dgv_Summary_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if ("5178.00081-加工折扣-(加工折扣)".Equals(dgv_Summary.Rows[e.RowIndex].HeaderCell.Value.ToString())) {
                    if (Convert.ToDecimal(dgv_Summary.Rows[e.RowIndex].Cells[e.ColumnIndex].Value) > 0) { 
                        MessageBox.Show("請輸入負數！");
                        dgv_Summary.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                    }
                    dgv_Summary.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.Red;
                }


                dgv_Summary.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = ConvertToDecimal(this.dgv_Summary.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);

                Summarise(e.RowIndex);
                /*
                if ("T0001".Equals(gTmplID))    // 製
                {
                    string bonus = "51301.0003-薪資支出-(直接人員年終獎金)";
                    string pension = "51301.0006-退休金-(直接人員退休金)";
                    string laborIsr = "51301.0004-保險費-(直接人員勞保)";
                    string HealthIsr = "51301.0005-保險費-(直接人員健保)";
                    string foodFee = "51301.0007-伙食費-(直接人員伙食費/每人每月)";

                    Calculate(bonus, pension, laborIsr, HealthIsr, foodFee, gDirectSalaryIndex, gDirectSalaryAddNumIndex, e.RowIndex);

                    string bonus2 = "5141-薪資支出-(間接人員年終獎金)";
                    string pension2 = "5178.0005-退休金-(間接人員退休金)";
                    string laborIsr2 = "5150-保險費-(間接人員勞保)";
                    string HealthIsr2 = "5150-保險費-(間接人員健保)";
                    string foodFee2 = "5156-伙食費-(間接人員伙食費/每人每月)";

                    Calculate(bonus2, pension2, laborIsr2, HealthIsr2, foodFee2, gIndirectSalaryIndex, gIndirectSalaryAddNumIndex, e.RowIndex);

                    
                }
                else if ("T0002".Equals(gTmplID))   // 銷
                {
                    string bonus = "6110-薪資支出-(年終獎金)";
                    string pension = "6188.0005-退休金-(退休金)";
                    string laborIsr = "6119-保險費-(勞保)";
                    string HealthIsr = "6119-保險費-(健保)";
                    string foodFee = "6127-伙食費-(伙食費/每人每月)";

                    Calculate(bonus, pension, laborIsr, HealthIsr, foodFee, gIndirectSalaryIndex, gIndirectSalaryAddNumIndex, e.RowIndex);

                    
                }
                else if ("T0003".Equals(gTmplID))   // 管
                {
                    string bonus = "6210-薪資支出-(年終獎金)";
                    string pension = "6288.0005-退休金-(退休金)";
                    string laborIsr = "6219-保險費-(勞保)";
                    string HealthIsr = "6219-保險費-(健保)";
                    string foodFee = "6227-伙食費-(伙食費/每人每月)";

                    Calculate(bonus, pension, laborIsr, HealthIsr, foodFee, gIndirectSalaryIndex, gIndirectSalaryAddNumIndex, e.RowIndex);

                    
                }
                else if ("T0004".Equals(gTmplID))   // 研
                {
                    string bonus = "6310-薪資支出-(年終獎金)";
                    string pension = "6388.0005-研-退休金-(退休金)";
                    string laborIsr = "6319-保險費-(勞保)";
                    string HealthIsr = "6319-保險費-(健保)";
                    string foodFee = "6327-伙食費-(伙食費/每人每月)";
                    Calculate(bonus, pension, laborIsr, HealthIsr, foodFee, gIndirectSalaryIndex, gIndirectSalaryAddNumIndex, e.RowIndex);
                }
                */

            }
            catch (Exception ex) {
                ExceptionHandle(dgv_Summary, e.RowIndex, e.ColumnIndex, ex.Message, "預算總表");
            }
        }

        private void Summarise(int RowIndex) {
            if ("T0001".Equals(gTmplID))    // 製
            {
                string bonus = "51301.0003-薪資支出-(直接人員年終獎金)";
                string pension = "51301.0006-退休金-(直接人員退休金)";
                string laborIsr = "51301.0004-保險費-(直接人員勞保)";
                string HealthIsr = "51301.0005-保險費-(直接人員健保)";
                string foodFee = "51301.0007-伙食費-(直接人員伙食費/每人每月)";

                Calculate(bonus, pension, laborIsr, HealthIsr, foodFee, gDirectSalaryIndex, gDirectSalaryAddNumIndex, RowIndex);

                string bonus2 = "5141-薪資支出-(間接人員年終獎金)";
                string pension2 = "5178.0005-退休金-(間接人員退休金)";
                string laborIsr2 = "5150-保險費-(間接人員勞保)";
                string HealthIsr2 = "5150-保險費-(間接人員健保)";
                string foodFee2 = "5156-伙食費-(間接人員伙食費/每人每月)";

                Calculate(bonus2, pension2, laborIsr2, HealthIsr2, foodFee2, gIndirectSalaryIndex, gIndirectSalaryAddNumIndex, RowIndex);
            }
            else if ("T0002".Equals(gTmplID))   // 銷
            {
                string bonus = "6110-薪資支出-(年終獎金)";
                string pension = "6188.0005-退休金-(退休金)";
                string laborIsr = "6119-保險費-(勞保)";
                string HealthIsr = "6119-保險費-(健保)";
                string foodFee = "6127-伙食費-(伙食費/每人每月)";

                Calculate(bonus, pension, laborIsr, HealthIsr, foodFee, gIndirectSalaryIndex, gIndirectSalaryAddNumIndex, RowIndex);


            }
            else if ("T0003".Equals(gTmplID))   // 管
            {
                string bonus = "6210-薪資支出-(年終獎金)";
                string pension = "6288.0005-退休金-(退休金)";
                string laborIsr = "6219-保險費-(勞保)";
                string HealthIsr = "6219-保險費-(健保)";
                string foodFee = "6227-伙食費-(伙食費/每人每月)";

                Calculate(bonus, pension, laborIsr, HealthIsr, foodFee, gIndirectSalaryIndex, gIndirectSalaryAddNumIndex, RowIndex);


            }
            else if ("T0004".Equals(gTmplID))   // 研
            {
                string bonus = "6310-薪資支出-(年終獎金)";
                string pension = "6388.0005-研-退休金-(退休金)";
                string laborIsr = "6319-保險費-(勞保)";
                string HealthIsr = "6319-保險費-(健保)";
                string foodFee = "6327-伙食費-(伙食費/每人每月)";
                Calculate(bonus, pension, laborIsr, HealthIsr, foodFee, gIndirectSalaryIndex, gIndirectSalaryAddNumIndex, RowIndex);
            }
        }



        /// <summary>
        /// 計算總表內，與薪資相關(年終、退休、勞健保、伙食費)的會計科目數值
        /// </summary>
        /// <param name="bonus">年終獎金會計科目完整名稱</param>
        /// <param name="pension">退休金會計科目完整名稱</param>
        /// <param name="laborIsr">勞保會計科目完整名稱</param>
        /// <param name="HealthIsr">健保會計科目完整名稱</param>
        /// <param name="foodFee">伙食費會計科目完整名稱</param>
        /// <param name="salaryIndex">[薪資]會計科目的RowIndex</param>
        /// <param name="salaryNumIndex">[薪資新增數]會計科目的RowIndex</param>
        /// <param name="rowIndex">目前正在編輯的欄位的RowIndex</param>
        
        private void Calculate(string bonus, string pension, string laborIsr, string HealthIsr, string foodFee, int salaryIndex, int salaryNumIndex, int rowIndex)
        {
            decimal v1 = 0;
            decimal tempSum = 0;    // 暫存特別科目的總計
            decimal itemSum = 0;    // 暫存各科目的總計
            //decimal monthlySum = 0; // 暫存各科目的各月份總計            
            
            for (int x = 0; x < dgv_Summary.RowCount - 1; x++)
            {                
                
                // 計算年終獎金
                if (bonus.Equals(dgv_Summary.Rows[x].HeaderCell.Value.ToString())) { 
                    tempSum = 0;
                    for (int y = 0; y < 12; y++)
                    {
                        // 加總「薪資支出-薪資 」和「薪資支出-薪資新增數」
                        v1 = Convert.ToDecimal(dgv_Summary.Rows[salaryIndex].Cells[y].Value) + Convert.ToDecimal(dgv_Summary.Rows[salaryNumIndex].Cells[y].Value);
                        dgv_Summary.Rows[x].Cells[y].Value = Math.Round((v1 / 12), 0, MidpointRounding.AwayFromZero);         // 年終獎金
                        //dgv_Summary.Rows[x].Cells[y].Value = ConvertToDecimal(dgv_Summary.Rows[x].Cells[y].Value);
                        tempSum += Convert.ToDecimal(dgv_Summary.Rows[x].Cells[y].Value);
                        dgv_Summary.Rows[x].Cells[12].Value = tempSum;

                    }
                }

                // 計算伙食費
                else if (foodFee.Equals(dgv_Summary.Rows[x].HeaderCell.Value.ToString())) { 
                    tempSum = 0;
                    for (int y = 0; y < 12; y++)
                    {
                        //dgv_Summary.Rows[x].Cells[y].Value = Math.Round(Convert.ToDecimal(dgv_Summary.Rows[x - 1].Cells[y].Value) * gFood, 0);         // 伙食費
                        dgv_Summary.Rows[x].Cells[y].Value = Convert.ToDecimal(dgv_Summary.Rows[x - 1].Cells[y].Value) * gFood;         // 伙食費
                        //dgv_Summary.Rows[x].Cells[y].Value = ConvertToDecimal(dgv_Summary.Rows[x].Cells[y].Value);
                        tempSum += Convert.ToDecimal(dgv_Summary.Rows[x].Cells[y].Value);
                        dgv_Summary.Rows[x].Cells[12].Value = tempSum;
                    }
                }

                // 計算退休金
                else if (pension.Equals(dgv_Summary.Rows[x].HeaderCell.Value.ToString())) { 
                    tempSum = 0;
                    for (int y = 0; y < 12; y++)
                    {
                        // 加總「薪資支出-薪資 」和「薪資支出-薪資新增數」
                        //v1 = Convert.ToDecimal(dgv_Summary.Rows[0].Cells[y].Value) + Convert.ToDecimal(dgv_Summary.Rows[1].Cells[y].Value);
                        v1 = Convert.ToDecimal(dgv_Summary.Rows[salaryIndex].Cells[y].Value) + Convert.ToDecimal(dgv_Summary.Rows[salaryNumIndex].Cells[y].Value);
                        dgv_Summary.Rows[x].Cells[y].Value = Math.Round((v1 * gPension), 0, MidpointRounding.AwayFromZero);   // 退休金
                        //dgv_Summary.Rows[x].Cells[y].Value = ConvertToDecimal(dgv_Summary.Rows[x].Cells[y].Value);
                        tempSum += Convert.ToDecimal(dgv_Summary.Rows[x].Cells[y].Value);
                        dgv_Summary.Rows[x].Cells[12].Value = tempSum;
                    }
                }

                // 計算勞保
                else if (laborIsr.Equals(dgv_Summary.Rows[x].HeaderCell.Value.ToString())) {
                    tempSum = 0;
                    for (int y = 0; y < 12; y++)
                    {
                        // 加總「薪資支出-薪資 」和「薪資支出-薪資新增數」
                        //v1 = Convert.ToDecimal(dgv_Summary.Rows[0].Cells[y].Value) + Convert.ToDecimal(dgv_Summary.Rows[1].Cells[y].Value);
                        v1 = Convert.ToDecimal(dgv_Summary.Rows[salaryIndex].Cells[y].Value) + Convert.ToDecimal(dgv_Summary.Rows[salaryNumIndex].Cells[y].Value);
                        dgv_Summary.Rows[x].Cells[y].Value = Math.Round((v1 * gLaborIsr), 0, MidpointRounding.AwayFromZero);  // 勞保
                        //dgv_Summary.Rows[x].Cells[y].Value = ConvertToDecimal(dgv_Summary.Rows[x].Cells[y].Value);
                        tempSum += Convert.ToDecimal(dgv_Summary.Rows[x].Cells[y].Value);
                        dgv_Summary.Rows[x].Cells[12].Value = tempSum;

                    }
                }

                // 計算健保
                else if (HealthIsr.Equals(dgv_Summary.Rows[x].HeaderCell.Value.ToString())) {
                    tempSum = 0;
                    for (int y = 0; y < 12; y++)
                    {
                        // 加總「薪資支出-薪資 」和「薪資支出-薪資新增數」
                        //v1 = Convert.ToDecimal(dgv_Summary.Rows[0].Cells[y].Value) + Convert.ToDecimal(dgv_Summary.Rows[1].Cells[y].Value);
                        v1 = Convert.ToDecimal(dgv_Summary.Rows[salaryIndex].Cells[y].Value) + Convert.ToDecimal(dgv_Summary.Rows[salaryNumIndex].Cells[y].Value);
                        dgv_Summary.Rows[x].Cells[y].Value = Math.Round((v1 * gHealthIsr), 0, MidpointRounding.AwayFromZero); // 健保
                        //dgv_Summary.Rows[x].Cells[y].Value = ConvertToDecimal(dgv_Summary.Rows[x].Cells[y].Value);
                        tempSum += Convert.ToDecimal(dgv_Summary.Rows[x].Cells[y].Value);
                        dgv_Summary.Rows[x].Cells[12].Value = tempSum;
                    }
                }

                //monthlySum += Convert.ToDecimal(dgv_Summary.Rows[x].Cells[clnIndex].Value);

                
            }

            //dgv_Summary.Rows[dgv_Summary.RowCount - 1].Cells[clnIndex].Value = monthlySum;
            //dgv_Summary.Rows[dgv_Summary.RowCount - 1].Cells[clnIndex].Value = ConvertToDecimal(dgv_Summary.Rows[dgv_Summary.RowCount - 1].Cells[clnIndex].Value);



            // 計算各項目總計
            if (rowIndex == -1) 
            {
                for (int j = 0; j < dgv_Summary.Rows.Count; j++) {
                    itemSum = 0;
                    for (int i = 0; i < 12; i++) { 
                        itemSum += Convert.ToDecimal(dgv_Summary.Rows[j].Cells[i].Value);
                        dgv_Summary.Rows[j].Cells[12].Value = itemSum;
                        dgv_Summary.Rows[j].Cells[12].Value = ConvertToDecimal(dgv_Summary.Rows[j].Cells[12].Value);
                    }                    
                }
            }
            else 
            { 
                for (int i = 0; i < 12; i++)
                    itemSum += Convert.ToDecimal(dgv_Summary.Rows[rowIndex].Cells[i].Value);
                dgv_Summary.Rows[rowIndex].Cells[12].Value = itemSum;
                dgv_Summary.Rows[rowIndex].Cells[12].Value = ConvertToDecimal(dgv_Summary.Rows[rowIndex].Cells[12].Value);
            }

            // 計算各月份總計
            ComputeMonthlySum(dgv_Summary, gMonthlySumList);
            
            /*
            gMonthlySumList.Clear();
            gMonthlySumList = ANBTI_Model.GetMonthlySum(dgv_Summary, gMonthlySumList);
            if (gMonthlySumList != null && gMonthlySumList.Count > 0) {
                for (int i = 0; i < 13; i++) {                      
                    dgv_Summary.Rows[dgv_Summary.RowCount - 1].Cells[i].Value = gMonthlySumList[i];
                    dgv_Summary.Rows[dgv_Summary.RowCount - 1].Cells[i].Value = ConvertToDecimal(dgv_Summary.Rows[dgv_Summary.RowCount - 1].Cells[i].Value);                    
                }
            }*/
        }

        /// <summary>
        /// 計算[預算總表]各月份(最後一筆的Row)的總計
        /// </summary>
        public void ComputeMonthlySum(DataGridView dgv, List<Object> MonthlySumList) {
            MonthlySumList.Clear();
            MonthlySumList = ANBTI_Model.GetMonthlySum(dgv, MonthlySumList);
            if (MonthlySumList != null && MonthlySumList.Count > 0)
            {
                for (int i = 0; i < 13; i++)
                {
                    dgv.Rows[dgv.RowCount - 1].Cells[i].Value = MonthlySumList[i];
                    dgv.Rows[dgv.RowCount - 1].Cells[i].Value = ConvertToDecimal(dgv.Rows[dgv.RowCount - 1].Cells[i].Value);
                }
            }
        }

        private void pbx_WriteToExcel_Click(object sender, EventArgs e)
        {
            _log.Debug("進入匯出Excel功能");
            StringBuilder sb = new StringBuilder();
            bool[] allTableState = null;


            if (RD_Depts.Contains(gMainDeptNo)) // 900新創、610開發
                allTableState = new bool[] { isANBTC_Saved, isANBTD_Saved, isANBTE_Saved, isANBTF_Saved, isANBTG_Saved, isANBTI_Saved, isANBTN_Saved };
            else
                allTableState = new bool[] { isANBTC_Saved, isANBTD_Saved, isANBTE_Saved, isANBTF_Saved, isANBTI_Saved, isANBTN_Saved };

            // 看看是否每張表都有儲存了
            for (int i = 0; i < allTableState.Count(); i++)
            {
                bool isTableSaved = allTableState[i];
                if (isTableSaved == false)
                {
                    sb.AppendLine("表-" + (i + 2) + " 尚未儲存");
                    _log.Debug("表-" + (i + 2) + " 尚未儲存");
                }
            }

            // 若有沒儲存的，先請user儲存
            if (sb.Length > 0)
            {
                sb.AppendLine();
                sb.Append("請將尚未儲存的表格儲存後，再匯出Excel");
                MessageBox.Show(sb.ToString());
                return;
            }



            DataGridView[] dgvs = new DataGridView[] {dgv_Org, dgv_HumanBudget, dgv_EmployeeTraining, dgv_BusinessTrip, dgv_CapEx, dgv_PRJ, dgv_PRJ_Content, dgv_Summary };

            string tmplFileName = "";

            if ("T0001".Equals(gTmplID))
                tmplFileName = "Templates_1.xlsx";
            else if ("T0002".Equals(gTmplID))
                tmplFileName = "Templates_2.xlsx";
            else if ("T0003".Equals(gTmplID))
                tmplFileName = "Templates_3.xlsx";
            else if ("T0004".Equals(gTmplID))
                tmplFileName = "Templates_4.xlsx";

            _log.Debug("使用者：" + gUserInfo.UserID + "，部門：" + gUserInfo.DeptNo + "，填寫部門代號為：" + gMainDeptNo + "的年度預算表。使用Excel樣版：" + tmplFileName);



            File_Model.WriteToExcel(gMainDeptNo, gMainDeptName, gUserInfo.DeptNo, cbx_Year.Text, tmplFileName, dgvs, keyValues, summaryList);

        }

        /// <summary>
        /// 建立總表內，自動帶出值的會計科目識別號
        /// </summary>
        private void CreateMap() {
            allMonthlyData.Clear();
            allMonthlyData.Add("1001", null);   //人力需求預算表-增減薪資金額-直接人員
            allMonthlyData.Add("1002", null);   //人力需求預算表-增減薪資金額-間接人員
            allMonthlyData.Add("2001", null);   //教育訓練計劃表-訓練費
            allMonthlyData.Add("3001", null);   //出差計劃表-差旅費
            allMonthlyData.Add("3002", null);   //出差計劃表-交際費
            allMonthlyData.Add("3003", null);   //出差計劃表-旅平險
            allMonthlyData.Add("4001", null);   //資本支出預算表-折舊
            allMonthlyData.Add("4002", null);   //資本支出預算表-攤銷
            allMonthlyData.Add("5001", null);   //研發專案彙總表-加班費
            allMonthlyData.Add("5002", null);   //研發專案彙總表-樣品費
            allMonthlyData.Add("5003", null);   //研發專案彙總表-技術移轉、設計費(勞務費)
            allMonthlyData.Add("5004", null);   //研發專案彙總表-認證費
            allMonthlyData.Add("5005", null);   //研發專案彙總表-模具(未滿8萬)及治具費
            allMonthlyData.Add("5006", null);   //研發專案彙總表-模具(8萬以上)
            allMonthlyData.Add("5007", null);   //研發專案彙總表-運費
            allMonthlyData.Add("6001", null);   //資本支出預算表-模具
            allMonthlyData.Add("8001", null);   //人力需求預算表-每月伙食費人數(間接)=實際人數+依起始月份增減之人數
            allMonthlyData.Add("8002", null);   //人力需求預算表-每月伙食費人數(直接)=實際人數+依起始月份增減之人數
        }

        /// <summary>
        /// 取得某些會計科目所需的費率
        /// </summary>
        private void GetRate() {
            gLaborIsr = ANBTJ_Model.GetRate("R0001");   // 勞保費率
            gHealthIsr = ANBTJ_Model.GetRate("R0002");  // 健保費率
            gFood = Math.Round(ANBTJ_Model.GetRate("R0003"), 0, MidpointRounding.AwayFromZero);       // 伙食費
            gPension = ANBTJ_Model.GetRate("R0004");    // 退休金提撥率
        }

        private void dgv_RD_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            decimal total = 0;
            decimal SMT_Pay = 0;    // 明躍支出
            decimal LH_Pay = 0;     // 立和支出
            decimal Cust_Pay = 0;   // 客戶支出
            try
            {
                // 轉換為decimal
                this.dgv_PRJ_Content.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = ConvertToDecimal(this.dgv_PRJ_Content.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);

                for (int i = 0; i < 12; i++)
                    if (this.dgv_PRJ_Content.Rows[e.RowIndex].Cells[i].Value != null && !String.IsNullOrEmpty(this.dgv_PRJ_Content.Rows[e.RowIndex].Cells[i].Value.ToString()))
                        total += Convert.ToDecimal(this.dgv_PRJ_Content.Rows[e.RowIndex].Cells[i].Value);   // 計算總金額


                this.dgv_PRJ_Content.Rows[e.RowIndex].Cells[12].Value = total;  // 總金額

                LH_Pay = Convert.ToDecimal(this.dgv_PRJ_Content.Rows[e.RowIndex].Cells[14].Value);     // 立和支出
                Cust_Pay = Convert.ToDecimal(this.dgv_PRJ_Content.Rows[e.RowIndex].Cells[15].Value);   // 客戶支出
                SMT_Pay = total - LH_Pay - Cust_Pay;     // 明躍支出 
                
                this.dgv_PRJ_Content.Rows[e.RowIndex].Cells[13].Value = SMT_Pay;
            }
            catch (Exception ex) {
                ExceptionHandle(dgv_PRJ_Content, e.RowIndex, e.ColumnIndex, ex.Message, "研發專案彙總表-單身");
            }


            //dgv_PRJ.Rows[dgv_PRJ.SelectedCells[0].RowIndex].HeaderCell.Value = null;
            //dgv_PRJ.Rows[dgv_PRJ.SelectedCells[0].RowIndex].HeaderCell.Style.BackColor = Color.Empty;

            dgv_PRJ.Rows[dgv_PRJ.SelectedCells[0].RowIndex].HeaderCell.Value = "V";     // 若單身有修改，則單頭的標頭也要標記為V
        }

        public decimal ConvertToDecimal(object obj) {
            bool result;
            string var = "";

            if (obj != null)
            {
                var = obj.ToString();
                result = isDecimal(var);

                if (result == true)
                    return Convert.ToDecimal(var);
                else
                    return 0;
            }
            return 0;
        }

        private void dgv_PRJ_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //dgv_PRJ.Rows[e.RowIndex].HeaderCell.Value = null;   // 有修改資料的話，標頭的儲存標記V先取消
            //dgv_PRJ.Rows[e.RowIndex].HeaderCell.Style.BackColor = Color.Empty;  //背景顏色也改成預設
            
            dgv_PRJ.Rows[e.RowIndex].HeaderCell.Value = "V";   // 有修改資料的話，標頭的儲存標記V
                                                               //dgv_PRJ.EnableHeadersVisualStyles = false;
                                                               //dgv_PRJ.Rows[e.RowIndex].HeaderCell.Style.BackColor = Color.LightSalmon;  //背景顏色也修改

            //dgv_PRJ.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

            /*
            try
            {
                for (int i = 0; i < dgv_PRJ.Rows.Count; i++)                
                    dgv_PRJ.Rows[i].Cells[0].Value = "RDP-" + (i + 1);
                
            }
            catch (Exception ex) {
                ExceptionHandle(dgv_PRJ, e.RowIndex, e.ColumnIndex, ex.Message, "研發專案彙總表-單頭");
            }*/
            CheckIsDate(e.RowIndex, dgv_PRJ.Rows[e.RowIndex].Cells["cln_StartPRJ"].Value, "cln_StartPRJ");
            CheckIsDate(e.RowIndex, dgv_PRJ.Rows[e.RowIndex].Cells["cln_EndPRJ"].Value, "cln_EndPRJ");
            CheckIsDate(e.RowIndex, dgv_PRJ.Rows[e.RowIndex].Cells["cln_TR_Day"].Value, "cln_TR_Day");
            
        }

        private bool CheckIsDate(int rowIndex, Object DateStr, string ClnName) {
            bool result1 = true;
            //bool result2 = true;
            string tempStr = (string)DateStr;
            if (DateStr != null && !String.IsNullOrEmpty(tempStr))
            {
                if (CheckIsDateFormat(tempStr) == true)
                {
                    string tempDateStr = tempStr.Clone().ToString();
                    if (tempDateStr.Length != 8 || tempDateStr.Contains("-") || tempDateStr.Contains("/"))
                    {
                        tempDateStr = tempDateStr.Replace("-", "");
                        tempDateStr = tempDateStr.Replace("/", "");
                    }

                    if(tempDateStr.Length != 8) {
                        MessageBox.Show("日期格式不正確，請重新檢查。日期格式必須為yyyyMMdd。例：20201027");
                        result1 = false;
                    }
                    else { 
                        dgv_PRJ.Rows[rowIndex].Cells[ClnName].Value = tempDateStr;
                        result1 =  true;
                    }
                }
                else
                    result1 = false;
            }
            else
                result1 = true;

            /*
            if (dgv_PRJ.Rows[rowIndex].Cells["cln_EndPRJ"].Value != null)
            {
                if (CheckIsDateFormat(dgv_PRJ.Rows[rowIndex].Cells["cln_EndPRJ"].Value.ToString()) == true)
                {
                    string dateStr = dgv_PRJ.Rows[rowIndex].Cells["cln_EndPRJ"].Value.ToString();
                    if (dateStr.Length != 8 || dateStr.Contains("-") || dateStr.Contains("/")) {
                        dateStr = dateStr.Replace("-", "");
                        dateStr = dateStr.Replace("/", "");
                    }

                    if (dateStr.Length != 8) {
                        MessageBox.Show("日期格式不正確，請重新檢查。日期格式必須為yyyyMMdd。例：20201027");
                        result2 = false;
                    }
                    else {
                        dgv_PRJ.Rows[rowIndex].Cells["cln_EndPRJ"].Value = dateStr;
                        result2 = true;
                    }
                }
                else
                    result2 = false;
            }
            else
                result2 = true;*/
            /*
            if (result1 == true && result2 == true)
                return true;
            else
                return false;    */
            return result1;
        }


        private bool CheckIsDateFormat(string dateStr) {
            string tempStr = dateStr.Clone().ToString();
            if (tempStr.Length == 8)
            {
                if (!tempStr.Contains("/") && !tempStr.Contains("-")) { 
                tempStr = tempStr.Substring(0, 4) + "/" +
                          tempStr.Substring(4, 2) + "/" +
                          tempStr.Substring(6, 2);
                }
            }

            DateTime dt = new DateTime();

            if (DateTime.TryParse(tempStr, out dt) == true) {
                return true;

            }
            else { 
                MessageBox.Show(dateStr + "日期格式不正確，請重新檢查。日期格式必須為yyyyMMdd。例：20201027");
                return false;
            }
        }


        /// <summary>
        /// 新增一筆空白單頭
        /// </summary>
        private void ADD_DGV_PRJ_Row() {

            // 專案序號EX: RDP-1, RDP-2...

            int indexNum = 0;   // 記錄目前的專案序號數字的部份
            int dashIndex = 0;  // 記錄專案序號「-」的位置
            string rdpid = "";
            try
            {
                if (dgv_PRJ.RowCount > 0) { 
                    dashIndex = dgv_PRJ.Rows[dgv_PRJ.Rows.Count - 1].Cells[0].Value.ToString().IndexOf("-");    // 記錄目前最後一筆的編號
                    indexNum = Convert.ToInt32(dgv_PRJ.Rows[dgv_PRJ.Rows.Count - 1].Cells[0].Value.ToString().Substring(dashIndex + 1));
                }

                dgv_PRJ.Rows.Add();     // 新增一筆ROW

                dgv_PRJ.Rows[dgv_PRJ.Rows.Count - 1].Cells[0].Value = "RDP-" + (indexNum + 1);  // 專案序號數字的部份 + 1

                rdpid = dgv_PRJ.Rows[dgv_PRJ.Rows.Count - 1].Cells[0].Value.ToString(); //記錄最後一筆的研發專案序號

                // 耐用年限預設值為4
                dgv_PRJ.Rows[dgv_PRJ.Rows.Count - 1].Cells[8].Value = 4;

                dgv_PRJ.RowHeadersWidth = 63;

                // 以單頭序號作為key值，單身物件暫且為null，放進MAP            
                keyValues.Add(rdpid, null);
            }
            catch (Exception ex) {
                _log.Debug("執行[新增專案]時，發生錯誤。" + ex.Message);
                MessageBox.Show("執行[新增專案]時，發生錯誤。" + ex.Message);
            }
        }

        /// <summary>
        /// 依據點選到的單頭號碼顯示單身的內容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_PRJ_Click(object sender, EventArgs e)
        {
            SwithchPRJ_Index();
            /*
            if (dgv_PRJ.RowCount > 0) {
                                
                if (gPRJ_Index == -1)   // 若為初始值-1
                {   
                    gPRJ_Index = dgv_PRJ.SelectedCells[0].RowIndex;     // 則記錄目前所選取的RowIndex
                    dgv_PRJ_Content = ANBTH_Model.SetDGV(dgv_PRJ, dgv_PRJ_Content, keyValues);
                }
                else
                {
                    if (gPRJ_Index != dgv_PRJ.SelectedCells[0].RowIndex) // 若上一筆跟這次點擊的是不同筆專案
                    {
                        if (dgv_PRJ.Rows[gPRJ_Index].HeaderCell.Value != null && "V".Equals(dgv_PRJ.Rows[gPRJ_Index].HeaderCell.Value))
                        {
                            //
                            if (MessageBox.Show("上一筆資料還沒儲存哦，請按「確定」儲存資料") == DialogResult.OK)
                            {
                                // 先儲存上一筆資料
                                SavePRJ_Data(gPRJ_Index);
                            //}
                            SavePRJ_Data(gPRJ_Index);   // 先儲存上一筆資料
                            lbl_PRJ_Notice.Text = "已編輯的專案序號 [" + dgv_PRJ.Rows[gPRJ_Index].Cells[0].Value + "] 儲存成功";
                            
                        }
                        //else
                        //{
                          //  gPRJ_Index = dgv_PRJ.SelectedCells[0].RowIndex;     // 記錄目前所選取的RowIndex                         
                        //}

                        dgv_PRJ_Content = ANBTH_Model.SetDGV(dgv_PRJ, dgv_PRJ_Content, keyValues);
                    }
                    else
                    {
                        //gPRJ_Index = dgv_PRJ.SelectedCells[0].RowIndex;     // 記錄目前所選取的RowIndex
                        //MessageBox.Show("目前選取到的RowIndex = " + gPRJ_Index);
                    }
                    gPRJ_Index = dgv_PRJ.SelectedCells[0].RowIndex;     // 記錄目前所選取的RowIndex
                }

                dgv_PRJ_Content.Visible = true;
            }*/
        }

        private void SwithchPRJ_Index() {
            if (dgv_PRJ.RowCount > 0)
            {

                if (gPRJ_Index == -1)   // 若為初始值-1
                {
                    gPRJ_Index = dgv_PRJ.SelectedCells[0].RowIndex;     // 則記錄目前所選取的RowIndex
                    dgv_PRJ_Content = ANBTH_Model.SetDGV(dgv_PRJ, dgv_PRJ_Content, keyValues);
                }
                else
                {
                    if (gPRJ_Index != dgv_PRJ.SelectedCells[0].RowIndex) // 若上一筆跟這次點擊的是不同筆專案
                    {
                        if (dgv_PRJ.Rows[gPRJ_Index].HeaderCell.Value != null && "V".Equals(dgv_PRJ.Rows[gPRJ_Index].HeaderCell.Value))
                        {
                            /*
                            if (MessageBox.Show("上一筆資料還沒儲存哦，請按「確定」儲存資料") == DialogResult.OK)
                            {
                                // 先儲存上一筆資料
                                SavePRJ_Data(gPRJ_Index);
                            }*/
                            SavePRJ_Data(gPRJ_Index);   // 先儲存上一筆資料
                            lbl_PRJ_Notice.Text = "已編輯的專案序號 [" + dgv_PRJ.Rows[gPRJ_Index].Cells[0].Value + "] 儲存成功";

                        }
                        //else
                        //{
                        //  gPRJ_Index = dgv_PRJ.SelectedCells[0].RowIndex;     // 記錄目前所選取的RowIndex                         
                        //}

                        dgv_PRJ_Content = ANBTH_Model.SetDGV(dgv_PRJ, dgv_PRJ_Content, keyValues);
                    }
                    else
                    {
                        //gPRJ_Index = dgv_PRJ.SelectedCells[0].RowIndex;     // 記錄目前所選取的RowIndex
                        //MessageBox.Show("目前選取到的RowIndex = " + gPRJ_Index);
                    }
                    gPRJ_Index = dgv_PRJ.SelectedCells[0].RowIndex;     // 記錄目前所選取的RowIndex
                }

                dgv_PRJ_Content.Visible = true;
            }
        }


        /// <summary>
        /// 若部門為610或900，則顯示研發專案彙總表
        /// </summary>
        private void SetTabControlAuth()
        {
            if (RD_Depts.Contains(gMainDeptNo))
            {
                PageG.Parent = null;
                PageF.Parent = tabControl1;
                PageG.Parent = tabControl1;
            }
            else
                PageF.Parent = null;
        }

        
        /// <summary>
        /// 載入各個DataGridView的內容
        /// </summary>
        /// <param name="annualBudgetFormID"></param>
        /// <param name="dgv"></param>
        public void Load_and_Set_DGV_Data(string annualBudgetFormID, DataGridView dgv, string DeptNo, string year, string TmplID, DataTable tmplTable, bool isDeptAccounting)
        {

            //List<Object> list = Common_Model.LoadMatrixData(annualBudgetFormID, gMainDeptNo, cbx_Year.Text, gTmplID);
            List<Object> list = Common_Model.LoadMatrixData(annualBudgetFormID, DeptNo, year, TmplID);


            if (list != null && list.Count > 0)
            {
                dgv = Common_Model.SetDGV(annualBudgetFormID, dgv, list, tmplTable, isDeptAccounting, DeptNo);
            }
            else
            {
                if (!"ABM007".Equals(annualBudgetFormID))   // 總表不能清空
                    dgv.Rows.Clear();
            }

            if (!"ABM007".Equals(annualBudgetFormID))
                SetRowsIndex(dgv);

            if ("ABM007".Equals(annualBudgetFormID))
                Set_dgvSummaryContent(dgv, tmplTable, isDeptAccounting);
        }

        /// <summary>
        /// 載入研發專案彙總表單頭單身的資料至DataGridView
        /// </summary>
        /// <param name="annualBudgetFormID"></param>
        /// <param name="dgv"></param>
        private void LoadAndSet_ANBTG_DGV(string annualBudgetFormID, DataGridView dgv_title, DataGridView dgv_content) {

            //Dictionary<string, Object> ANBTG_kvs = ANBTG_Model.LoadData(annualBudgetFormID, gUserInfo.DeptNo, cbx_Year.Text);
            Dictionary<string, Object> ANBTG_kvs = ANBTG_Model.LoadData(annualBudgetFormID, cbx_Dept.Text, cbx_Year.Text);
            //Dictionary<string, Object> ANBTH_kVs = null;

            if (ANBTG_kvs != null && ANBTG_kvs.Count > 0)
            {
                ANBTG_kvs = ANBTH_Model.LoadData(ANBTG_kvs); // 取得單身資料
                keyValues = ANBTG_kvs;  // 將單身資料指派給全域的map

                dgv_title = ANBTG_Model.SetDGV(dgv_title, ANBTG_kvs);
            }
            else {
                dgv_title.Rows.Clear();
                dgv_content.Visible = false;
            }

            if (dgv_PRJ.RowCount < 1)
                btn_SavePRJ.Enabled = false;
            else
                btn_SavePRJ.Enabled = true;
        }

        /// <summary>
        /// 允許特定部門開啟特定按鈕
        /// </summary>
        public void OpenDeptButton() {

            if (gUserInfo.ParentDept != null)
            {
                if ("820".Equals(gUserInfo.ParentDept.DeptNo) || "820".Equals(gUserInfo.DeptNo) ||    // 財會部
                    "110".Equals(gUserInfo.ParentDept.DeptNo) || "110".Equals(gUserInfo.DeptNo))      // 稽核室(暫時代理))  
                {      
                    btn_OpenRateForm.Enabled = true;
                    btn_OpenTmplDeptRef_Form.Enabled = true;
                    btn_forAccounting.Visible = true;

                }
            }
            else if (gUserInfo.ParentDept != null &&  "510".Equals(gUserInfo.DeptNo))   // 510製造部，顯示「510匯報表」按鍵            
            {
                btn_Dept5_Rpt.Visible = true;                
            }

        }


        /// <summary>
        /// 鎖住和開放特定部門對於人力需求預算表的直接人員和間接人員選項
        /// </summary>
        private void lockHumanType() {
            
            string[] unlockDepts = new string[] {"512", "514", "515" };
            DataGridViewComboBoxColumn cbc = (DataGridViewComboBoxColumn)dgv_HumanBudget.Columns["dgv_HBT_cln_Role"];
            cbc.Items.Clear();
            if (unlockDepts.Contains(gMainDeptNo)){
                cbc.Items.Add("直接人員");
                cbc.Items.Add("間接人員");
            }
            else            
                cbc.Items.Add("間接人員");            
        }


        private void btn_OpenRateForm_Click(object sender, EventArgs e)
        {
            Form_Rate form = new Form_Rate();
            form.gUserId = gUserInfo.UserID;
            form.gDeptNo = gUserInfo.DeptNo;
            OpenForm(form);
        }

        private void btn_OpenTmplDeptRef_Form_Click(object sender, EventArgs e)
        {
            Form_Dept_Tmpl_Ref form = new Form_Dept_Tmpl_Ref();
            OpenForm(form);
        }

        /// <summary>
        /// 自動帶入每個表格的序號
        /// </summary>
        private void setTableNum() {
            for (int i = 0; i < tabControl1.TabPages.Count; i++) {
                TabPage page = tabControl1.TabPages[i];
                if (page.Parent != null) 
                {
                    if (page.Text.Contains("-")) {
                        page.Text = page.Text.Substring(page.Text.IndexOf("-") + 1, page.Text.Length - page.Text.IndexOf("-") -1 );
                        page.Text = "表" + (i + 1) + "-" + page.Text;
                    }
                    else
                        page.Text = "表" + (i + 1) + "-" + page.Text;
                }
            }
        }

        /// <summary>
        /// 統一處理編輯DataGridView時發生的Exception
        /// </summary>
        /// <param name="dgv">目標DataGridView</param>
        /// <param name="rowIndex">目標的RowIndex</param>
        /// <param name="columnIndex">目標的ColumnIndex</param>
        /// <param name="errMsg">錯誤訊息</param>
        /// <param name="tableName">正在編輯的報表名稱</param>
        private void ExceptionHandle(DataGridView dgv, int rowIndex, int columnIndex, string errMsg, string tableName) {
            string v = "";
            if (dgv.Rows[rowIndex].Cells[columnIndex].Value != null)
                v = dgv.Rows[rowIndex].Cells[columnIndex].Value.ToString();
            else
                v = "null";

            _log.Debug("編輯[" + tableName + "]時發生錯誤。Row: " + rowIndex + "、Column: " + columnIndex + "、Value: " + v);
            _log.Debug(errMsg);
            MessageBox.Show("編輯[" + tableName + "]時發生錯誤。Row: " + rowIndex + "、Column: " + columnIndex + "、Value: " + v + "\n" + errMsg);
        }

        private void cbx_Dept_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*
            gMainDeptNo = cbx_Dept.Text;
            Load_and_Set_DGV_Data("ABM001", dgv_HumanBudget, gMainDeptNo, cbx_Year.Text, gTmplID, null);
            Load_and_Set_DGV_Data("ABM002", dgv_EmployeeTraining, gMainDeptNo, cbx_Year.Text, gTmplID, null);
            Load_and_Set_DGV_Data("ABM003", dgv_BusinessTrip, gMainDeptNo, cbx_Year.Text, gTmplID, null);
            Load_and_Set_DGV_Data("ABM004", dgv_CapEx, gMainDeptNo, cbx_Year.Text, gTmplID, null);
            LoadAndSet_ANBTG_DGV("ABM005", dgv_PRJ, dgv_PRJ_Content);
            Load_and_Set_DGV_Data("ABM007", dgv_Summary, gMainDeptNo, cbx_Year.Text, gTmplID, gDT_Summary);
            */
        }

        private Dictionary<string, Object> RemoveDeletedObject(Dictionary<string, Object> keyValues) {
            int objectCount = keyValues.Count;
            Dictionary<string, Object> tempKVs = new Dictionary<string, Object>();


            if (keyValues != null && keyValues.Count > 0) {

                // 先把目前的keyValues複製一份到另一個Dictionary
                for (int i = 0; i < objectCount; i++) {
                    string key = keyValues.ElementAt(i).Key;
                    Object obj = keyValues.ElementAt(i).Value;
                    tempKVs.Add(key, obj);
                }


                for (int i = 0; i < objectCount; i++) {
                    ANBTG anbtg = (ANBTG)tempKVs.ElementAt(i).Value;

                    if ("Y".Equals(anbtg.Tg013))
                        keyValues.Remove(anbtg.Tg004);
                }
            }

            return keyValues;
        }

        /// <summary>
        /// 暫存研發專案彙總表的資料
        /// </summary>
        private void SavePRJ_Data(int rowIndex) {
            string annualBudgetFormID = "ABM005";
            if (dgv_PRJ.Rows.Count >= 1)
            {
                //dgv_PRJ_Content.Visible = true;

                keyValues = ANBTG_Model.SaveMatrixData(keyValues, dgv_PRJ_Content, dgv_PRJ.Rows[rowIndex], annualBudgetFormID, gMainDeptNo, gUserInfo.DeptNo, cbx_Year.Text, gUserInfo.UserID, gToday);

                //dgv_PRJ.Rows[dgv_PRJ.SelectedCells[0].RowIndex].HeaderCell.Value = "V";     // 有儲存過的row，在header標記 V
                //dgv_PRJ.EnableHeadersVisualStyles = false;
                //dgv_PRJ.Rows[dgv_PRJ.SelectedCells[0].RowIndex].HeaderCell.Style.BackColor = Color.Aqua;   // 並把顏色改成藍色

                dgv_PRJ.Rows[rowIndex].HeaderCell.Value = null;     // 有儲存過的row，header取消標記

                //dgv_PRJ.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;


                //MessageBox.Show("專案序號 [ " + dgv_PRJ.Rows[rowIndex].Cells[0].Value.ToString() + " ] 儲存成功");
            }
            else
                MessageBox.Show("目前沒有資料可供暫存");
        }

        private void btn_forAccounting_Click(object sender, EventArgs e)
        {
            Form_Accounting form = new Form_Accounting();

            form.gUserInfo = gUserInfo;
            form.gYear = cbx_Year.Text;

            OpenForm(form);
        }

        /// <summary>
        /// 取得全域的樣版編號
        /// </summary>
        private void GetTmplID() {

            if (gDT_Summary != null && gDT_Summary.Rows.Count > 0) {
                for (int i = 0; i < gDT_Summary.Rows.Count; i++) {
                    gTmplID = gDT_Summary.Rows[i]["TL002"].ToString();
                    if (!String.IsNullOrEmpty(gTmplID))
                        break;
                }
            }
        }

        private void btn_SwitchDept_Click(object sender, EventArgs e)
        {
            gMainDeptNo = cbx_Dept.Text;
            gPRJ_Index = -1; // 設回預設值
            /*
            Load_and_Set_DGV_Data("ABM001", dgv_HumanBudget, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
            Load_and_Set_DGV_Data("ABM002", dgv_EmployeeTraining, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
            Load_and_Set_DGV_Data("ABM003", dgv_BusinessTrip, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
            Load_and_Set_DGV_Data("ABM004", dgv_CapEx, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
            LoadAndSet_ANBTG_DGV("ABM005", dgv_PRJ, dgv_PRJ_Content);
            Load_and_Set_DGV_Data("ABM007", dgv_Summary, gMainDeptNo, cbx_Year.Text, gTmplID, gDT_Summary, false);
            Load_and_Set_DGV_Data("ABM008", dgv_Org, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
            */
            Init_Data();
            /*
            bool[] allTableState;
            if (RD_Depts.Contains(gMainDeptNo)) // 900新創、610開發
                allTableState = new bool[] { isANBTC_Saved, isANBTD_Saved, isANBTE_Saved, isANBTF_Saved, isANBTG_Saved, isANBTI_Saved, isANBTN_Saved };
            else
                allTableState = new bool[] { isANBTC_Saved, isANBTD_Saved, isANBTE_Saved, isANBTF_Saved, isANBTI_Saved, isANBTN_Saved };
            
            for (int i = 0; i < allTableState.Length; i++) {
                bool state = allTableState[i];
                state = false;
                allTableState[i] = state;
            }
            */

            ClearAllSaveState();

            for (int i = 0; i < tabControl1.TabPages.Count; i++)
                if (tabControl1.TabPages[i].Text.Contains("✔")) {
                    tabControl1.TabPages[i].Text = tabControl1.TabPages[i].Text.Substring(1, tabControl1.TabPages[i].Text.Length - 1);
                }
            _log.Debug("切換到部門" + gMainDeptNo);
            MessageBox.Show("已成功切換到部門" + gMainDeptNo);
            
        }

        private void btn_Dept5_Rpt_Click(object sender, EventArgs e)
        {
            Form_forDept5 form = new Form_forDept5();
            form.gYear = cbx_Year.Text;
            form.gDept = gMainDeptNo;
            OpenForm(form);
        }

        

        private void btn_OpenExcel_Click(object sender, EventArgs e)
        {            
            Process.Start(@"\\192.168.123.50\共用暫存區\IT\程式\!內部製作小程式\財會年度預算編列程式\編制時程與編制說明.xlsx");
        }

        
        private void btn_LastYear_Click(object sender, EventArgs e)
        {
            // 此功能未測試，暫不使用!!!

            string lastYear = (Convert.ToInt32(cbx_Year.Text) - 1).ToString();
            Load_and_Set_DGV_Data("ABM001", dgv_HumanBudget, gMainDeptNo, lastYear, gTmplID, null, false);
            ANBTC_Model.SetSyncRowColor(dgv_HumanBudget);

            Load_and_Set_DGV_Data("ABM002", dgv_EmployeeTraining, gMainDeptNo, lastYear, gTmplID, null, false);
            Load_and_Set_DGV_Data("ABM003", dgv_BusinessTrip, gMainDeptNo, lastYear, gTmplID, null, false);
            Load_and_Set_DGV_Data("ABM004", dgv_CapEx, gMainDeptNo, lastYear, gTmplID, null, false);
            LoadAndSet_ANBTG_DGV("ABM005", dgv_PRJ, dgv_PRJ_Content);
            Load_and_Set_DGV_Data("ABM007", dgv_Summary, gMainDeptNo, lastYear, gTmplID, gDT_Summary, false);
        }

        /// <summary>
        /// 搜尋資料並載入到主畫面
        /// </summary>
        private void Init_Data() {
            lbl_PRJ_Notice.Text = "資料全部暫存後，" + Environment.NewLine + "記得按下方[儲存本表]，" + Environment.NewLine + "才能正確儲存至資料庫";
            List<Object> allDeptList = new List<object>();
            SetYear();  // 設定年份
            
            /*
            bool isManager = false;
            
            if (gUserInfo.ParentDept != null &&
                ("820".Equals(gUserInfo.ParentDept.DeptNo) || "820".Equals(gUserInfo.DeptNo)) ||    // 財會部
                ("110".Equals(gUserInfo.ParentDept.DeptNo) || "110".Equals(gUserInfo.DeptNo))       // 稽核室(暫時代理)
                )
            {
                DataTable dt = ANBTK_Model.LoadDepts(cbx_Year.Text);
                gList_depts.Clear();
                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)                    
                        gList_depts.Add(dt.Rows[i]["TK002"]);                    
                }
                
                allDeptList.Clear();
                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                        allDeptList.Add(dt.Rows[i]["TK002"]);
                }
                isManager = true;
            }
            else { 
                gList_depts = ANBTK_Model.LoadDepts(gUserInfo, cbx_Year.Text, gList_depts);
                isManager = false;
            }*/

            DataTable dt = ANBTK_Model.LoadDepts(cbx_Year.Text);

            allDeptList.Clear();
            if (dt != null && dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                    allDeptList.Add(dt.Rows[i]["TK002"]);   // 取得所有部門列表
            }
            gList_depts.Clear();
            gList_depts = ANBTK_Model.LoadDepts(gUserInfo, cbx_Year.Text, gList_depts);  // 取得自已所屬部門列表


            
            try
            {
                if (gList_depts != null && gList_depts.Count > 0)
                {
                    cbx_Dept.Items.Clear();
                    if (gUserInfo.ParentDept != null)
                    {
                        if (("820".Equals(gUserInfo.ParentDept.DeptNo) || "820".Equals(gUserInfo.DeptNo)) ||    // 財會部
                            ("110".Equals(gUserInfo.ParentDept.DeptNo) || "110".Equals(gUserInfo.DeptNo)))      // 稽核室(暫時代理))
                        {
                            for (int i = 0; i < allDeptList.Count; i++)
                                cbx_Dept.Items.Add(allDeptList[i]);

                            string myDept = gList_depts[0].ToString();

                            for (int i = 0; i < cbx_Dept.Items.Count; i++)
                            {
                                if (String.IsNullOrEmpty(gMainDeptNo))
                                {
                                    if (myDept.Equals(cbx_Dept.Items[i].ToString()))
                                    {
                                        cbx_Dept.SelectedItem = cbx_Dept.Items[i];
                                        break;
                                    }
                                }
                                else
                                {
                                    if (gMainDeptNo.Equals(cbx_Dept.Items[i].ToString()))
                                    {
                                        cbx_Dept.SelectedItem = cbx_Dept.Items[i];
                                        break;
                                    }
                                }
                            }
                        }
                        else {
                            for (int i = 0; i < gList_depts.Count; i++) { 
                                cbx_Dept.Items.Add(gList_depts[i]);
                            }

                            // 新部門910的職務代理人的工號
                            if ("20191202".Equals(gUserInfo.UserID))
                            {
                                // 加入被職代的部門代號
                                cbx_Dept.Items.Add("910");
                            }




                            if (String.IsNullOrEmpty(gMainDeptNo))
                                cbx_Dept.SelectedIndex = 0;
                            else
                            {
                                for (int i = 0; i < cbx_Dept.Items.Count; i++)
                                {
                                    if (gMainDeptNo.Equals(cbx_Dept.Items[i].ToString()))
                                    {
                                        cbx_Dept.SelectedIndex = i;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < gList_depts.Count; i++)
                            cbx_Dept.Items.Add(gList_depts[i]);

                        // 職務代理人
                        if ("Y".Equals(config.AppSettings.Settings["isSub"].Value.ToString()))
                        {
                            // Sub_Stuff_ID 職務代理人的工號
                            if (config.AppSettings.Settings["Sub_Staff_ID"].Value.ToString().Equals(gUserInfo.UserID))
                            {
                                // 加入被職代的部門代號
                                cbx_Dept.Items.Add(config.AppSettings.Settings["SubDept"].Value.ToString());
                            }
                        }
    

                        if (String.IsNullOrEmpty(gMainDeptNo))
                            cbx_Dept.SelectedIndex = 0;
                        else
                        {
                            for (int i = 0; i < cbx_Dept.Items.Count; i++) {
                                if (gMainDeptNo.Equals(cbx_Dept.Items[i].ToString()))
                                {
                                    cbx_Dept.SelectedIndex = i;
                                }
                            }   
                        }                        
                    }

                    gMainDeptNo = cbx_Dept.SelectedItem.ToString();


                    if (cbx_Dept.Items.Count > 1)
                        btn_SwitchDept.Enabled = true;
                    else
                        btn_SwitchDept.Enabled = false;


                    //gMainDeptNo = cbx_Dept.Text;
                    //if (String.IsNullOrEmpty(gMainDeptNo))
                    //    gMainDeptNo = cbx_Dept.Items[0].ToString();

                    gDT_Summary = Tmpl_Model.GetTmplContentByDept(gMainDeptNo);

                    GetTmplID();    // 取得全域的樣版編號

                    //dgv_Summary = Set_dgvSummary(gMainDeptNo, dgv_Summary, gDT_Summary);
                    dgv_Summary = ANBTI_Model.Set_dgvSummary(gDT_Summary, dgv_Summary, false, gMainDeptNo);
                }
            

                //if (cbx_Dept.Items.Count > 0)
                //    cbx_Dept.SelectedIndex = 0;

                _log.Debug("填寫部門代號為：" + gMainDeptNo + "的年度預算表");

                if (gMainDeptNo.Equals(gUserInfo.DeptNo))
                    gMainDeptName = gUserInfo.DeptName;
                else if (gUserInfo.ParentDept != null && gMainDeptNo.Equals(gUserInfo.ParentDept.DeptNo))
                    gMainDeptName = gUserInfo.ParentDept.DeptName;

                if (String.IsNullOrEmpty(gMainDeptNo))
                {
                    MessageBox.Show("您不需填寫年度預算表，或沒有填寫的權限。\n" + "請按確定關閉程式");
                    this.Close();
                }
                else
                {
                    ClearAllDGV();
                    lockHumanType();
                    //Load_and_Set_DGV_Data(string annualBudgetFormID, DataGridView dgv, string DeptNo, string year, string TmplID)
                    //List<Object> list = Common_Model.LoadMatrixData(annualBudgetFormID, gMainDeptNo, cbx_Year.Text, gTmplID);                    

                    Load_and_Set_DGV_Data("ABM001", dgv_HumanBudget, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
                    ANBTC_Model.SetSyncRowColor(dgv_HumanBudget);

                    Load_and_Set_DGV_Data("ABM002", dgv_EmployeeTraining, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
                    Load_and_Set_DGV_Data("ABM003", dgv_BusinessTrip, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
                    Load_and_Set_DGV_Data("ABM004", dgv_CapEx, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
                    LoadAndSet_ANBTG_DGV("ABM005", dgv_PRJ, dgv_PRJ_Content);
                    Load_and_Set_DGV_Data("ABM007", dgv_Summary, gMainDeptNo, cbx_Year.Text, gTmplID, gDT_Summary, false);
                    Load_and_Set_DGV_Data("ABM008", dgv_Org, gMainDeptNo, cbx_Year.Text, gTmplID, null, false);
                    
                    Form_FeeFillIn.LoadOtherFee(cbx_Dept.Text, "F001", "", dgv_Summary);       // 載入瓦斯費
                    Form_FeeFillIn.LoadOtherFee(cbx_Dept.Text, "F002", "F003", dgv_Summary);   // 載入水電費
                    Form_FeeFillIn.LoadOtherFee(cbx_Dept.Text, "F004", "", dgv_Summary);       // 載入電話費
                    Form_FeeFillIn.LoadOtherFee(cbx_Dept.Text, "F005", "", dgv_Summary);       // 載入網路費
                    Form_FeeFillIn.LoadOtherFee(cbx_Dept.Text, "F006", "", dgv_Summary);       // 載入折舊-現有固定資產
                    Form_FeeFillIn.LoadOtherFee(cbx_Dept.Text, "F007", "", dgv_Summary);       // 載入各項攤提-現有無形資產
                    Form_FeeFillIn.LoadOtherFee(cbx_Dept.Text, "F008", "", dgv_Summary);       // 載入雜費-預付費用(維護合約等)攤提



                    Set_dgvPRJ_Content();

                    CreateMap();    // 建立總表內，自動帶出值的會科識別號
                    GetRate();      // 取得某些會科所需的費率
                    SetTabControlAuth();

                    //tbx_MainDeptNo.Text = gMainDeptNo;

                    if (gUserInfo.ParentDept != null)
                    {
                        tbx_ParentDeptNo.Text = gUserInfo.ParentDept.DeptNo;
                        tbx_ParentDeptName.Text = gUserInfo.ParentDept.DeptName;
                    }

                    tbx_DeptNo.Text = gUserInfo.DeptNo;
                    tbx_DeptName.Text = gUserInfo.DeptName;

                    setTableNum();

                    OpenDeptButton();
                }
            }
            catch (Exception ex)
            {
                _log.Debug("資料初始化載入時發生錯誤：" + ex.Message);
                MessageBox.Show("資料初始化載入時發生錯誤：" + ex.Message);
            }
        }

        private void ClearAllDGV() 
        {
            DataGridView[] dgvs = new DataGridView[]{ dgv_Org , dgv_HumanBudget };
            for (int i = 0; i < dgvs.Length; i++) {
                DataGridView dgv = dgvs[i];
                dgv.Rows.Clear();
            }
        }

        private void ClearAllSaveState() {
            isANBTC_Saved = false;  // [人力需求預算表]的儲存狀態
            isANBTD_Saved = false;  // [教育訓練計劃表]的儲存狀態
            isANBTE_Saved = false;  // [出差計劃表]的儲存狀態
            isANBTF_Saved = false;  // [資本支出預算表]的儲存狀態
            isANBTG_Saved = false;  // [研發專案彙總表]的儲存狀態
            isANBTI_Saved = false;  // [預算總表]的儲存狀態
            isANBTN_Saved = false;  // [組織編制表]的儲存狀態
        }
        
    }
}