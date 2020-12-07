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
    public partial class Form_FeePreFillIn : Form
    {
        public UserInfo gUserInfo = new UserInfo();
        private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Form_FeePreFillIn()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
        }

        private void Form_FeePreFillIn_Load(object sender, EventArgs e)
        {

            Dictionary<string, DataGridView> keyValues = new Dictionary<string, DataGridView>();
            keyValues.Add("F001", dgv_FPF_1);
            keyValues.Add("F002", dgv_FPF_2);
            keyValues.Add("F003", dgv_FPF_3);
            keyValues.Add("F004", dgv_FPF_4);
            keyValues.Add("F005", dgv_FPF_5);


            DataTable Depts = LoadDepts();  // 載入部門資料

            InitDGVs(Depts,  keyValues);    // 初始化DataGridView

            Load_and_SetDGV();  // 資料庫讀取資料並設置到DataGridView


        }

        /// <summary>
        /// 載入部門資料
        /// </summary>
        /// <returns></returns>
        private DataTable LoadDepts() {
            DataTable dt_DeptNo = ANBTK_Model.LoadDepts();
            DataTable dt_DeptDetail = Tmpl_Model.GetDept_Table();

            if (dt_DeptNo != null && dt_DeptNo.Rows.Count > 0) {
                dt_DeptNo.Columns.Add("部門名稱");
            }

            for (int i = 0; i < dt_DeptNo.Rows.Count; i++) {
                string select = "部門代號 = " + dt_DeptNo.Rows[i]["部門代號"].ToString();
                DataRow[] rows = dt_DeptDetail.Select(select);

                if (rows != null && rows.Count() > 0) {
                    dt_DeptNo.Rows[i]["部門名稱"] = rows[0]["部門名稱"].ToString();
                }
            }

            return dt_DeptNo;
        }

        /// <summary>
        /// 初始化DataGridView
        /// </summary>
        /// <param name="dt_Depts"></param>
        /// <param name="dgvs_Map"></param>
        private void InitDGVs(DataTable dt_Depts, Dictionary<string, DataGridView> dgvs_Map) {
            if (dgvs_Map != null && dgvs_Map.Count() > 0)
            {
                for (int x = 0; x < dgvs_Map.Count(); x++)
                {
                    DataGridView dgv = dgvs_Map.Values.ElementAt(x);
                    dgv.Rows.Clear();

                    // 第一列
                    dgv.Rows.Add();
                    dgv.Rows[0].HeaderCell.Value = "000";
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        if (dgv.Columns[j].Name.Contains("cln_DeptName"))
                        {
                            dgv.Rows[0].Cells[j].Value = "每月預算";                            
                        }
                        else if (dgv.Columns[j].Name.Contains("cln_FeeType"))
                        {
                            dgv.Rows[0].Cells[j].Value = dgvs_Map.Keys.ElementAt(x);                           
                        }
                        
                    }
                    dgv.Rows[0].Cells[2].ReadOnly = true;
                    dgv.Rows[0].DefaultCellStyle.BackColor = Color.LightSalmon;


                    // 第二列到最後一最
                    
                    for (int i = 0; i < dt_Depts.Rows.Count; i++)
                    {
                        dgv.Rows.Add();

                        dgv.Rows[i + 1].HeaderCell.Value = dt_Depts.Rows[i]["部門代號"].ToString();                        
                        for (int j = 0; j < dgv.Columns.Count; j++)
                        {
                            if (dgv.Columns[j].Name.Contains("cln_DeptName"))                                                           
                                dgv.Rows[i + 1].Cells[j].Value = dt_Depts.Rows[i]["部門名稱"].ToString();
                                                            
                            else if (dgv.Columns[j].Name.Contains("cln_FeeType"))                            
                                dgv.Rows[i + 1].Cells[j].Value = dgvs_Map.Keys.ElementAt(x);
                        }
                    }
                    dgv.RowHeadersWidth = 70;
                }
            }
        }

        private void Load_and_SetDGV() 
        {
            // 載入ANBTM資料並放到DataGridView中
            List<Object> F001_List = ANBTM_Model.Load_ANBTM("F001", "");
            SetDGV(F001_List, dgv_FPF_1);
            tbx_FPF_1.Text = GetSumRate(dgv_FPF_1).ToString();
            ChangeBackColor(tbx_FPF_1);

            List<Object> F002_List = ANBTM_Model.Load_ANBTM("F002", "");
            SetDGV(F002_List, dgv_FPF_2);
            tbx_FPF_2.Text = GetSumRate(dgv_FPF_2).ToString();
            ChangeBackColor(tbx_FPF_2);

            List<Object> F003_List = ANBTM_Model.Load_ANBTM("F003", "");
            SetDGV(F003_List, dgv_FPF_3);
            tbx_FPF_3.Text = GetSumRate(dgv_FPF_3).ToString();
            ChangeBackColor(tbx_FPF_3);

            List<Object> F004_List = ANBTM_Model.Load_ANBTM("F004", "");
            SetDGV(F004_List, dgv_FPF_4);
            tbx_FPF_4.Text = GetSumRate(dgv_FPF_4).ToString();
            ChangeBackColor(tbx_FPF_4);

            List<Object> F005_List = ANBTM_Model.Load_ANBTM("F005", "");
            SetDGV(F005_List, dgv_FPF_5);
            tbx_FPF_5.Text = GetSumRate(dgv_FPF_5).ToString();
            ChangeBackColor(tbx_FPF_5);
        }

        private void dgv_FPF_1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //decimal sumRate = 100;
            ComputeSharePrice(dgv_FPF_1);
            tbx_FPF_1.Text = GetSumRate(dgv_FPF_1).ToString();
            ChangeBackColor(tbx_FPF_1);
            if (100 == Convert.ToInt32(tbx_FPF_1.Text))
                GetBalance(dgv_FPF_1);
        }
        private void dgv_FPF_2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ComputeSharePrice(dgv_FPF_2);
            tbx_FPF_2.Text = GetSumRate(dgv_FPF_2).ToString();
            ChangeBackColor(tbx_FPF_2);
            if (100 == Convert.ToInt32(tbx_FPF_2.Text))
                GetBalance(dgv_FPF_2);
        }
        private void dgv_FPF_3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ComputeSharePrice(dgv_FPF_3);
            tbx_FPF_3.Text = GetSumRate(dgv_FPF_3).ToString();
            ChangeBackColor(tbx_FPF_3);
        }
        private void dgv_FPF_4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ComputeSharePrice(dgv_FPF_4);
            tbx_FPF_4.Text = GetSumRate(dgv_FPF_4).ToString();
            ChangeBackColor(tbx_FPF_4);
        }

        private void dgv_FPF_5_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ComputeSharePrice(dgv_FPF_5);
            tbx_FPF_5.Text = GetSumRate(dgv_FPF_5).ToString();
            ChangeBackColor(tbx_FPF_5);
        }

        private void ComputeSharePrice(DataGridView dgv) 
        {
            decimal balance = 0;
            if (dgv != null && dgv.RowCount > 0) 
            {   
                for (int i = 0; i < dgv.RowCount; i++) // 列
                {
                    decimal SumOfDept = 0;
                    decimal rate = 0;
                    if (i >= 1) {
                        rate = Convert.ToDecimal(dgv.Rows[i].Cells[2].Value) / 100;   // 比率                        
                    }

                    for (int j = 3; j < dgv.Columns.Count - 1; j++) // 欄
                    {
                        if (i >= 1) { 
                            decimal budgetPerM = Convert.ToDecimal(dgv.Rows[0].Cells[j].Value); // 每月預算
                            dgv.Rows[i].Cells[j].Value = Math.Round(rate * budgetPerM, 0, MidpointRounding.AwayFromZero);
                            dgv.Rows[i].Cells[j].ReadOnly = true;

                            balance = budgetPerM - Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        }

                        SumOfDept += Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);  
                    
                    }
                    // 最後一欄合計
                    dgv.Rows[i].Cells[dgv.Columns.Count - 1].Value = SumOfDept;
                }
            }
        }

        private void GetBalance(DataGridView dgv)
        {
            decimal sum = 0;
            decimal balance = 0;
            if (dgv != null && dgv.RowCount > 0)
            {
                for (int i = 3; i < dgv.Columns.Count - 1; i++) // 欄
                {                    
                    for (int j = 0; j < dgv.RowCount - 1; j++) // 列
                    {
                        
                        if (j == 0)
                            balance = Convert.ToDecimal(dgv.Rows[0].Cells[i].Value); // 每月預算

                        else if (j >= 1)
                            balance -= Convert.ToDecimal(dgv.Rows[j].Cells[i].Value);
                        

                        //SumOfDept += Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);

                    }
                    // 最後一欄合計
                    dgv.Rows[dgv.RowCount - 1].Cells[i].Value = balance;
                    sum += balance;
                }
                dgv.Rows[dgv.RowCount - 1].Cells[dgv.Columns.Count - 1].Value = sum;
            }
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            bool result;

            List<Object> list1 = DgvToList(dgv_FPF_1, gUserInfo.UserID, gUserInfo.DeptNo);
            if ((bool)list1[0] == true)
                result = ANBTM_Model.UpdateToANBTM((List<Object>)list1[1]);
            else { 
                MessageBox.Show("瓦斯費比率總合不為100，請重新檢查。");
                return;
            }


            // 儲存DataGridView2              
            List<Object> list2 = DgvToList(dgv_FPF_2, gUserInfo.UserID, gUserInfo.DeptNo);
            if ((bool)list2[0] == true)
                result = ANBTM_Model.UpdateToANBTM((List<Object>)list2[1]);
            else { 
                MessageBox.Show("水費比率總合不為100，請重新檢查。");
                return;
            }

            // 儲存DataGridView3            
            List<Object> list3 = DgvToList(dgv_FPF_3, gUserInfo.UserID, gUserInfo.DeptNo);
            if ((bool)list3[0] == true)
                result = ANBTM_Model.UpdateToANBTM((List<Object>)list3[1]);
            else { 
                MessageBox.Show("電費比率總合不為100，請重新檢查。");
                return;
            }

            // 儲存DataGridView4
            List<Object> list4 = DgvToList(dgv_FPF_4, gUserInfo.UserID, gUserInfo.DeptNo);
            if ((bool)list4[0] == true)
                result = ANBTM_Model.UpdateToANBTM((List<Object>)list4[1]);
            else { 
                MessageBox.Show("電話費比率總合不為100，請重新檢查。");
                return;
            }

            // 儲存DataGridView5            
            List<Object> list5 = DgvToList(dgv_FPF_5, gUserInfo.UserID, gUserInfo.DeptNo);
            if ((bool)list5[0] == true)
                result = ANBTM_Model.UpdateToANBTM((List<Object>)list5[1]);
            else { 
                MessageBox.Show("網路費比率總合不為100，請重新檢查。");
                return;
            }


            // 儲存完畢後，重新載入資料
            Load_and_SetDGV();

            MessageBox.Show("儲存完畢！");
        }

        private List<Object> DgvToList(DataGridView dgv, string userId, string deptNo) {
            List<Object> result = new List<object>();
            List<Object> list = new List<Object>();
            decimal SumOfRate = 0;
            if (dgv != null && dgv.RowCount > 0) 
            {                

                for (int i = 0; i < dgv.RowCount; i++) 
                {
                    ANBTM anbtm = new ANBTM();

                    anbtm.Creator = userId;     // 資料建立者
                    anbtm.Usr_group = deptNo;   // 部門代號
                    anbtm.Create_date = DateTime.Now.ToString("yyyyMMdd");

                    anbtm.Modifier = userId;     // 資料建立者
                    anbtm.Modi_date = DateTime.Now.ToString("yyyyMMdd");
                    

                    anbtm.Tm001 = dgv.Rows[i].HeaderCell.Value.ToString();

                    for (int j = 0; j < dgv.Columns.Count; j++) {
                        if (dgv.Columns[j].Name.Contains("DeptName"))
                            anbtm.Tm002 = dgv.Rows[i].Cells[j].Value.ToString();
                        else if (dgv.Columns[j].Name.Contains("Rate")) { 
                            anbtm.Tm003 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                            SumOfRate += anbtm.Tm003;
                        }
                        else if (dgv.Columns[j].Name.Contains("M01"))
                            anbtm.Tm004 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("M02"))
                            anbtm.Tm005 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("M03"))
                            anbtm.Tm006 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("M04"))
                            anbtm.Tm007 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("M05"))
                            anbtm.Tm008 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("M06"))
                            anbtm.Tm009 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("M07"))
                            anbtm.Tm010 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("M08"))
                            anbtm.Tm011 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("M09"))
                            anbtm.Tm012 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("M10"))
                            anbtm.Tm013 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("M11"))
                            anbtm.Tm014 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("M12"))
                            anbtm.Tm015 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("Sum"))
                            anbtm.Tm016 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Columns[j].Name.Contains("FeeType"))
                            anbtm.Tm017 = dgv.Rows[i].Cells[j].Value.ToString();
                        /*if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("DeptName"))
                            anbtm.Tm002 = dgv.Rows[i].Cells[j].Value.ToString();
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("Rate"))
                            anbtm.Tm003 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M01"))
                            anbtm.Tm004 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M02"))
                            anbtm.Tm005 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M03"))
                            anbtm.Tm006 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M04"))
                            anbtm.Tm007 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M05"))
                            anbtm.Tm008 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M06"))
                            anbtm.Tm009 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M07"))
                            anbtm.Tm010 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M08"))
                            anbtm.Tm011 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M09"))
                            anbtm.Tm012 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M10"))
                            anbtm.Tm013 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M11"))
                            anbtm.Tm014 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("M12"))
                            anbtm.Tm015 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("Sum"))
                            anbtm.Tm016 = Convert.ToDecimal(dgv.Rows[i].Cells[j].Value);
                        else if (dgv.Rows[i].Cells[j].OwningColumn.Name.Contains("FeeType"))
                            anbtm.Tm017 = dgv.Rows[i].Cells[j].Value.ToString();*/
                    }

                    switch (anbtm.Tm017) {
                        case "F001":
                            anbtm.Tm018 = "瓦斯";
                            break;
                        case "F002":
                            anbtm.Tm018 = "水費";
                            break;
                        case "F003":
                            anbtm.Tm018 = "電費";
                            break;
                        case "F004":
                            anbtm.Tm018 = "電話費";
                            break;
                        case "F005":
                            anbtm.Tm018 = "網路費";
                            break;                        
                    }

                    list.Add(anbtm);
                }                
            }
            if (SumOfRate == 100)
                result.Add(true);
            else
                result.Add(false);

            result.Add(list);
                

            return result;
        }


        private void SetDGV(List<Object> list, DataGridView dgv) {
            if (list != null && list.Count > 0) {
                for (int i = 0; i < list.Count; i++) { 
                    ANBTM anbtm = (ANBTM)list[i];

                    for (int j = 0; j < dgv.Rows.Count; j++) 
                    {
                        if (anbtm.Tm001.Equals(dgv.Rows[j].HeaderCell.Value.ToString()))
                        {
                            for (int x = 0; x < dgv.Columns.Count; x++) {
                            
                                if (dgv.Columns[x].Name.Contains("DeptName"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm002;
                                else if (dgv.Columns[x].Name.Contains("Rate"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm003;
                                else if (dgv.Columns[x].Name.Contains("M01"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm004;
                                else if (dgv.Columns[x].Name.Contains("M02"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm005;
                                else if (dgv.Columns[x].Name.Contains("M03"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm006;
                                else if (dgv.Columns[x].Name.Contains("M04"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm007;
                                else if (dgv.Columns[x].Name.Contains("M05"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm008;
                                else if (dgv.Columns[x].Name.Contains("M06"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm009;
                                else if (dgv.Columns[x].Name.Contains("M07"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm010;
                                else if (dgv.Columns[x].Name.Contains("M08"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm011;
                                else if (dgv.Columns[x].Name.Contains("M09"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm012;
                                else if (dgv.Columns[x].Name.Contains("M10"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm013;
                                else if (dgv.Columns[x].Name.Contains("M11"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm014;
                                else if (dgv.Columns[x].Name.Contains("M12"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm015;
                                else if (dgv.Columns[x].Name.Contains("Sum"))
                                    dgv.Rows[j].Cells[x].Value = anbtm.Tm016;
                                
                            }
                            break;
                        }
                    }
                }
            }
        }


        private decimal GetSumRate(DataGridView dgv) {
            decimal result = 0;

            for (int i = 1; i < dgv.RowCount; i++) {
                result += Convert.ToDecimal(dgv.Rows[i].Cells[2].Value);
            }

            return Math.Round(result, 0, MidpointRounding.AwayFromZero);        
        }

        private void ChangeBackColor(TextBox tbx) {
            if ("100".Equals(tbx.Text))
                tbx.BackColor = Color.White;
            else
                tbx.BackColor = Color.LightCoral;
        }
        
    }
}
