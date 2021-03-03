using AnnualBudget.BOs;
using AnnualBudget.Model;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AnnualBudget
{
    public partial class Form_FeeFillIn : Form
    {
        private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        Form_Main gForm_Main = new Form_Main();
        ANBTK anbtk = null;
        public UserInfo gUserInfo = new UserInfo();
        DataTable gDeptsTable = null;                    // 儲存需填寫年度預算表部門的DataTable
        DataTable gTmplTable = new DataTable();     // 儲存樣版的DataTable
        public List<Object> gMonthlySumList = new List<object>();

        public Form_FeeFillIn()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
        }

        private void Form_FeeFillIn_Load(object sender, EventArgs e)
        {
            SetYear();
            LoadDeptData(cbx_Year.Text);
            
        }

        /// <summary>
        ///  設定起始年份
        /// </summary>
        public void SetYear()
        {
            int year = DateTime.Today.Year;

            for (int i = year; i < year + 3; i++)
            {
                cbx_Year.Items.Add(i);
            }

            cbx_Year.SelectedIndex = 1;
        }

        public void LoadDeptData(string year) {
            gDeptsTable = ANBTK_Model.LoadDepts(year);
            
            if (gDeptsTable != null && gDeptsTable.Rows.Count > 0) {
                for (int i = 0; i < gDeptsTable.Rows.Count; i++) {
                   cbx_Dept.Items.Add(gDeptsTable.Rows[i]["TK002"]);
                }
            }
        }
                

        private void dgv_Summary_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dgv_Summary.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = gForm_Main.ConvertToDecimal(dgv_Summary.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);

            decimal itemSum = 0;    // 暫存各科目的總計
            for (int i = 0; i < 12; i++)
                itemSum += Convert.ToDecimal(dgv_Summary.Rows[e.RowIndex].Cells[i].Value);
            dgv_Summary.Rows[e.RowIndex].Cells[12].Value = itemSum;
            dgv_Summary.Rows[e.RowIndex].Cells[12].Value = gForm_Main.ConvertToDecimal(dgv_Summary.Rows[e.RowIndex].Cells[12].Value);
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            gForm_Main.ComputeMonthlySum(dgv_Summary, gMonthlySumList);
            List<Object> myList = ANBTI_Model.SaveMatrixDataForAccounting(dgv_Summary, gTmplTable, "ABM007", cbx_Dept.Text, gUserInfo.DeptNo, cbx_Year.Text, gUserInfo.UserID, DateTime.Now.ToString("yyyyMMdd"));

            if (ANBTI_Model.UpdateToANBTI(myList) == false)
            {
                _log.Debug("財會儲存「會科項目預填」至DB時發生錯誤！");
                MessageBox.Show("財會儲存「會科項目預填」至DB時發生錯誤！");
            }
            else
            {                
                _log.Debug("儲存[會科項目預填]成功！重新載入資料");
                gForm_Main.Load_and_Set_DGV_Data("ABM007", dgv_Summary, cbx_Dept.Text, cbx_Year.Text, anbtk.Tk004, gTmplTable, true);      // 將資料重新載入
                MessageBox.Show("儲存[會科項目預填]成功！");
            }

        }

        public static void LoadOtherFee(string dept, string feeCode1, string feeCode2, DataGridView dgv)
        {
            List<Object> list = ANBTM_Model.GetOtherFee(dept, feeCode1, feeCode2);

            if (list != null && list.Count > 0)
            {
                for (int i = 0; i < list.Count; i++)
                {

                    ANBTM anbtm = (ANBTM)list[i];

                    for (int x = 0; x < dgv.RowCount; x++)
                    {
                        // 比對名稱
                        if (dgv.Rows[x].Visible == true && anbtm.Tm017.Equals(dgv.Rows[x].Tag))
                        {

                            for (int y = 0; y < dgv.ColumnCount; y++)
                            {
                                if (dgv.Columns[y].Name.Contains("M01"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm004;
                                else if (dgv.Columns[y].Name.Contains("M02"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm005;
                                else if (dgv.Columns[y].Name.Contains("M03"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm006;
                                else if (dgv.Columns[y].Name.Contains("M04"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm007;
                                else if (dgv.Columns[y].Name.Contains("M05"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm008;
                                else if (dgv.Columns[y].Name.Contains("M06"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm009;
                                else if (dgv.Columns[y].Name.Contains("M07"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm010;
                                else if (dgv.Columns[y].Name.Contains("M08"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm011;
                                else if (dgv.Columns[y].Name.Contains("M09"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm012;
                                else if (dgv.Columns[y].Name.Contains("M10"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm013;
                                else if (dgv.Columns[y].Name.Contains("M11"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm014;
                                else if (dgv.Columns[y].Name.Contains("M12"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm015;
                                else if (dgv.Columns[y].Name.Contains("Sum_Item"))
                                    dgv.Rows[x].Cells[y].Value = anbtm.Tm016;

                            }
                        }
                    }
                }
            }
        }

        private void btn_Search_Click(object sender, EventArgs e)
        {
            _log.Debug("按下搜尋");


            dgv_Summary.Rows.Clear();
            if (!String.IsNullOrEmpty(cbx_Dept.Text) && !String.IsNullOrEmpty(cbx_Year.Text))
            {
                string whereStr = "TK002 = " + cbx_Dept.Text;

                gTmplTable = Tmpl_Model.GetTmplContentByDept(cbx_Dept.Text);   // 取得樣版內容

                //dgv_Summary = gForm_Main.Set_dgvSummary(cbx_Depts.Text, dgv_Summary, gTmplTable);

                // 將樣版套用到DataGridView上，設定總表的列的RowHeaderCell
                dgv_Summary = ANBTI_Model.Set_dgvSummary(gTmplTable, dgv_Summary, true, gUserInfo.DeptNo, false);

                DataRow[] rows = gDeptsTable.Select(whereStr);

                if (rows != null && rows.Count() > 0)
                {
                    anbtk = new ANBTK();
                    anbtk.Tk001 = rows[0][0].ToString();    // 唯一序號
                    anbtk.Tk002 = rows[0][1].ToString();    // 部門代號
                    anbtk.Tk003 = rows[0][2].ToString();    // 年度
                    anbtk.Tk004 = rows[0][3].ToString();    // 樣版編號
                    anbtk.Tk005 = rows[0][4].ToString();    // 樣版版本號
                }


                gForm_Main.Load_and_Set_DGV_Data("ABM007", dgv_Summary, cbx_Dept.Text, cbx_Year.Text, anbtk.Tk004, gTmplTable, true);

                LoadOtherFee(cbx_Dept.Text, "F001", "", dgv_Summary);       // 載入瓦斯費
                LoadOtherFee(cbx_Dept.Text, "F002", "F003", dgv_Summary);   // 載入水電費
                LoadOtherFee(cbx_Dept.Text, "F004", "", dgv_Summary);       // 載入電話費
                LoadOtherFee(cbx_Dept.Text, "F005", "", dgv_Summary);       // 載入網路費
                LoadOtherFee(cbx_Dept.Text, "F006", "", dgv_Summary);       // 載入折舊-現有固定資產
                LoadOtherFee(cbx_Dept.Text, "F007", "", dgv_Summary);       // 載入各項攤提-現有無形資產
                LoadOtherFee(cbx_Dept.Text, "F008", "", dgv_Summary);       // 載入雜費-預付費用(維護合約等)攤提

            }
            else
            {
                MessageBox.Show("部門與年度不得為空值");
            }
        }
    }
}
