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
    public partial class Form_forDept5 : Form
    {
        private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public string gYear = "";
        public string gDept = "";
        public List<Object> summaryList = new List<Object>();
        DataTable table = null;


        public Form_forDept5()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
        }

        private void Form_forDept5_Load(object sender, EventArgs e)
        {
            LoadData();
        }


        private void LoadData() {
            table = ANBTI_Model.LoadDept5_SumData(gYear);


            if (table != null && table.Rows.Count > 0)
            {
                dgv_Summary = ANBTI_Model.Set_dgvSummary(table, dgv_Summary, false, gDept);

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    dgv_Summary.Rows[i].Cells[0].Value = Convert.ToDecimal(table.Rows[i]["TI010"]);
                    dgv_Summary.Rows[i].Cells[1].Value = Convert.ToDecimal(table.Rows[i]["TI011"]);
                    dgv_Summary.Rows[i].Cells[2].Value = Convert.ToDecimal(table.Rows[i]["TI012"]);
                    dgv_Summary.Rows[i].Cells[3].Value = Convert.ToDecimal(table.Rows[i]["TI013"]);
                    dgv_Summary.Rows[i].Cells[4].Value = Convert.ToDecimal(table.Rows[i]["TI014"]);
                    dgv_Summary.Rows[i].Cells[5].Value = Convert.ToDecimal(table.Rows[i]["TI015"]);
                    dgv_Summary.Rows[i].Cells[6].Value = Convert.ToDecimal(table.Rows[i]["TI016"]);
                    dgv_Summary.Rows[i].Cells[7].Value = Convert.ToDecimal(table.Rows[i]["TI017"]);
                    dgv_Summary.Rows[i].Cells[8].Value = Convert.ToDecimal(table.Rows[i]["TI018"]);
                    dgv_Summary.Rows[i].Cells[9].Value = Convert.ToDecimal(table.Rows[i]["TI019"]);
                    dgv_Summary.Rows[i].Cells[10].Value = Convert.ToDecimal(table.Rows[i]["TI020"]);
                    dgv_Summary.Rows[i].Cells[11].Value = Convert.ToDecimal(table.Rows[i]["TI021"]);
                    dgv_Summary.Rows[i].Cells[12].Value = Convert.ToDecimal(table.Rows[i]["TI022"]);                    
                }
            }
        }

        private List<Object> ConvertToList(DataTable table) {
            List<Object> myList = null;

            if (table != null && table.Rows.Count > 0) {
                myList = new List<object>();

                for (int i = 0; i < table.Rows.Count; i++) {
                    ANBTI anbti = new ANBTI();
                    anbti.Ti003 = table.Rows[i]["TI003"].ToString();
                    anbti.Ti006 = Convert.ToDecimal(table.Rows[i]["TI006"]);
                    anbti.Ti007 = table.Rows[i]["會科編號"].ToString();
                    anbti.Ti008 = table.Rows[i]["會科名稱"].ToString();
                    anbti.Ti009 = table.Rows[i]["會科說明"].ToString();
                    anbti.Ti010 = Convert.ToDecimal(table.Rows[i]["TI010"]);
                    anbti.Ti011 = Convert.ToDecimal(table.Rows[i]["TI011"]);
                    anbti.Ti012 = Convert.ToDecimal(table.Rows[i]["TI012"]);
                    anbti.Ti013 = Convert.ToDecimal(table.Rows[i]["TI013"]);
                    anbti.Ti014 = Convert.ToDecimal(table.Rows[i]["TI014"]);
                    anbti.Ti015 = Convert.ToDecimal(table.Rows[i]["TI015"]);
                    anbti.Ti016 = Convert.ToDecimal(table.Rows[i]["TI016"]);
                    anbti.Ti017 = Convert.ToDecimal(table.Rows[i]["TI017"]);
                    anbti.Ti018 = Convert.ToDecimal(table.Rows[i]["TI018"]);
                    anbti.Ti019 = Convert.ToDecimal(table.Rows[i]["TI019"]);
                    anbti.Ti020 = Convert.ToDecimal(table.Rows[i]["TI020"]);
                    anbti.Ti021 = Convert.ToDecimal(table.Rows[i]["TI021"]);
                    anbti.Ti022 = Convert.ToDecimal(table.Rows[i]["TI022"]);
                   

                    myList.Add(anbti);
                }
            }
            return myList;
            
        }

        private void pbx_WriteToExcel_Click(object sender, EventArgs e)
        {
            _log.Debug("進入匯出Excel功能");
            //StringBuilder sb = new StringBuilder();
            //bool[] allTableState = null;
            string templateExcelName = "Templates_5.xlsx";

            string departNo = "510";
            string departName = "製造部";

            using (FileStream fsIn = new FileStream("Templates\\" + templateExcelName, FileMode.Open))
            {
                _log.Debug("檔案路徑：" + "Templates\\" + templateExcelName);

                //開啟Excel
                XSSFWorkbook workbook = new XSSFWorkbook(fsIn);

                fsIn.Close();


                summaryList = ConvertToList(table);


                /*
                if (RD_Depts.Contains(gMainDeptNo)) // 900新創、610開發
                    allTableState = new bool[] { isANBTC_Saved, isANBTD_Saved, isANBTE_Saved, isANBTF_Saved, isANBTG_Saved, isANBTI_Saved };
                else
                    allTableState = new bool[] { isANBTC_Saved, isANBTD_Saved, isANBTE_Saved, isANBTF_Saved, isANBTI_Saved };

                // 看看是否每張表都有儲存了
                for (int i = 0; i < allTableState.Count(); i++)
                {
                    bool isTableSaved = allTableState[i];
                    if (isTableSaved == false)
                    {
                        sb.AppendLine("表-" + (i + 1) + " 尚未儲存");
                        _log.Debug("表-" + (i + 1) + " 尚未儲存");
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
                */


                DataGridView[] dgvs = new DataGridView[] { dgv_Summary };

                //string tmplFileName = "";
                /*
                if ("T0001".Equals(gTmplID))
                    tmplFileName = "Templates_1.xlsx";
                else if ("T0002".Equals(gTmplID))
                    tmplFileName = "Templates_2.xlsx";
                else if ("T0003".Equals(gTmplID))
                    tmplFileName = "Templates_3.xlsx";
                else if ("T0004".Equals(gTmplID))
                    tmplFileName = "Templates_4.xlsx";

                _log.Debug("使用者：" + gUserInfo.UserID + "，部門：" + gUserInfo.DeptNo + "，填寫部門代號為：" + gMainDeptNo + "的年度預算表。使用Excel樣版：" + tmplFileName);
                */


                //File_Model.WriteToExcel(gMainDeptNo, gMainDeptName, gYear, "Templates_5.xlsx", dgvs, keyValues, summaryList);

                _log.Debug("匯出預算總表...");
                ISheet sheet4 = workbook.GetSheetAt(4); // 預算總表
                sheet4 = File_Model.WriteSummaryToExcel(departNo, departName, gDept, gYear, sheet4, dgvs[0], summaryList);


                SaveFileDialog dlg = new SaveFileDialog();
                dlg.FileName = gYear + "年部門預算表-(" + departNo + ").xlsx"; // Default file name                
                dlg.Filter = "Excel活頁簿 | *.xlsx"; // Filter files by extension
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
    }
}
