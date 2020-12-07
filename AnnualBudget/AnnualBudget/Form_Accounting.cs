using AnnualBudget.BOs;
using AnnualBudget.Model;
using log4net.Config;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AnnualBudget
{
    public partial class Form_Accounting : Form
    {
        public UserInfo gUserInfo = new UserInfo();
        public string gYear = "";
        public string gDept = "";
        private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Form_Accounting()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
        }
                

        public DialogResult OpenForm(Form form)
        {
            form.ShowDialog();

            return form.DialogResult;
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
            form.gUserInfo = gUserInfo;
            OpenForm(form);
        }

        private void btn_Dept5_Rpt_Click(object sender, EventArgs e)
        {
            Form_forDept5 form = new Form_forDept5();
            form.gYear = gYear;
            OpenForm(form);
        }

        private void btn_FeeFillIn_Click(object sender, EventArgs e)
        {
            Form_FeeFillIn form = new Form_FeeFillIn();
            form.gUserInfo = gUserInfo;
            
            OpenForm(form);
        }

        private void btn_GetSum_of_AccountsAndDepts_Click(object sender, EventArgs e)
        {
            DataTable dt_Dept = ANBTI_Model.GetSumByDept(gYear);
            DataTable dt_Accounts = ANBTI_Model.GetSumByAccounts(gYear);

            _log.Debug("進入匯出Excel功能");
            
            string templateExcelName = "Templates_6.xlsx";
                        
            using (FileStream fsIn = new FileStream("Templates\\" + templateExcelName, FileMode.Open))
            {
                _log.Debug("檔案路徑：" + "Templates\\" + templateExcelName);

                //開啟Excel
                XSSFWorkbook workbook = new XSSFWorkbook(fsIn);

                fsIn.Close();

                _log.Debug("匯出各科目及各部門預算總計...");
                ISheet sheet0 = workbook.GetSheetAt(0); // 各科目總計
                sheet0 = File_Model.WriteSumByAccountsToExcel(gYear, sheet0, dt_Accounts);
                ISheet sheet1 = workbook.GetSheetAt(1); // 各部門預算總計
                sheet1 = File_Model.WriteSumByDeptsToExcel(gYear, sheet1, dt_Dept);


                SaveFileDialog dlg = new SaveFileDialog();
                dlg.FileName = gYear + "年各科目及各部門預算總計.xlsx";    // Default file name                
                dlg.Filter = "Excel活頁簿 | *.xlsx";                     // Filter files by extension
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    //另存為Result.xls
                    using (FileStream fsOut = new FileStream(dlg.FileName, FileMode.Create))
                    {
                        workbook.Write(fsOut);
                        fsOut.Close();
                    }
                }
            }
        }

        private void btn_FeePreFillIn_Click(object sender, EventArgs e)
        {
            Form_FeePreFillIn form = new Form_FeePreFillIn();
            form.gUserInfo = gUserInfo;

            OpenForm(form);
        }

        private void btn_AuthMgmt_Click(object sender, EventArgs e)
        {
            Form_AuthMgmt form = new Form_AuthMgmt();
            form.gUserInfo = gUserInfo;

            OpenForm(form);
        }
    }
}
