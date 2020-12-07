using AnnualBudget.BOs;
using AnnualBudget.Model;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AnnualBudget
{
    public partial class Form_Dept_Tmpl_Ref : Form
    {
        Dictionary<string, string> tmpl_No_Name_pairs = new Dictionary<string, string>();
        public UserInfo gUserInfo = new UserInfo();

        public Form_Dept_Tmpl_Ref()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
        }

        private void Form_Dept_Tmpl_Ref_Load(object sender, EventArgs e)
        {
            SetDeptList();  // 抓出公司部門列表
            SetTmpl_ID();   // 抓出所有樣版編號
            SetYear();      // 設定年份
        }

        private void cbx_Tmpl_ID_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetTmplVer(cbx_Tmpl_ID.SelectedItem.ToString());
        }

        /// <summary>
        /// 抓出公司部門列表
        /// </summary>
        public void SetDeptList() {
            DataTable table = Tmpl_Model.GetDept_Table();
            dgv_Dept.DataSource = table;
        }

        /// <summary>
        /// 透過樣版編號，取得使用該樣版的部門列表
        /// </summary>
        /// <param name="tmplID"></param>
        /// <param name="ver"></param>
        public void Get_Tmpl_Dept_Ref_Table(string tmplID, string ver)
        {
            DataTable table2 = Tmpl_Model.GetDept_Tmpl_Ref_Table(tmplID, ver);

            /*
            // 先幫DataGridView建立與table2相同的Row數
            if (table2 != null && table2.Rows.Count > 0) {
                int dgvRowsCount = dgv_Tmpl_Dep_Ref.RowCount;

                if (table2.Rows.Count > dgvRowsCount) {
                    for (int i = 0; i < table2.Rows.Count - dgvRowsCount; i++) {
                        dgv_Tmpl_Dep_Ref.Rows.Add();
                    }
                }
            }

            // 增加完Row數後，把資料塞到DataGridView裡
            for (int i = 0; i < table2.Rows.Count; i++) {
                dgv_Tmpl_Dep_Ref.Rows[i].Cells[0].Value = table2.Rows[i]["部門代號"];
            }
            */

            dgv_Tmpl_Dep_Ref.DataSource = table2;
        }

        /// <summary>
        /// 抓出所有樣版編號
        /// </summary>
        public void SetTmpl_ID() {
            DataTable dt = Tmpl_Model.GetTmplID();
            
            if (dt != null && dt.Rows.Count > 0) { 
                for (int i = 0; i < dt.Rows.Count; i++) { 
                    cbx_Tmpl_ID.Items.Add(dt.Rows[i][0]);
                    tmpl_No_Name_pairs.Add(dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString());
                }
            }
        }
                

        public void SetTmplVer(string tmplID)
        {
            // 取得版本號
            DataTable dt = Tmpl_Model.GetTmplVer(tmplID);

            cbx_Ver.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)            
                cbx_Ver.Items.Add(dt.Rows[i][0]);

            cbx_Ver.SelectedIndex = 0;


            // 設定模版名稱
            tbx_TmplName.Text = tmpl_No_Name_pairs[tmplID];
        }

        

        private void btn_SearchTmpl_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(cbx_Tmpl_ID.Text) && !String.IsNullOrEmpty(cbx_Ver.Text))
            {
                string tmplID = cbx_Tmpl_ID.SelectedItem.ToString();
                string ver = cbx_Ver.SelectedItem.ToString();
                SetTmplContent(tmplID, ver);            // 取得樣版內容
                Get_Tmpl_Dept_Ref_Table(tmplID, ver);   // 透過樣版編號，取得使用該樣版的部門列表
                btn_Add.Enabled = true;
                btn_Remove.Enabled = true;
            }
            else
            {
                MessageBox.Show("「樣版編號」或「樣版版本號」未選擇內容。");
            }
        }

        /// <summary>
        /// 取得樣版內容
        /// </summary>
        /// <param name="tmplID"></param>
        /// <param name="ver"></param>
        public void SetTmplContent(string tmplID, string ver) {
            DataTable dt = Tmpl_Model.GetTmplContent(tmplID, ver);
            dgv_Tmpl.DataSource = dt;

            int NotVisibledColumnsCount = dgv_Tmpl.ColumnCount - 2;

            for (int i = 0; i < dgv_Tmpl.ColumnCount; i++) {
                dgv_Tmpl.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
                if (i >= NotVisibledColumnsCount)
                    dgv_Tmpl.Columns[i].Visible = false;
            }
        }

        private void btn_Add_Click(object sender, EventArgs e)
        {
            AddRow((DataTable)dgv_Tmpl_Dep_Ref.DataSource);
        }

        private void btn_Remove_Click(object sender, EventArgs e)
        {
            RemoveRow((DataTable)dgv_Tmpl_Dep_Ref.DataSource);
        }

        public void AddRow(DataTable dt) {
            DataRow[] drs = dt.Select(dt.Columns[0].ColumnName + " = " + dgv_Dept.SelectedRows[0].Cells[0].Value.ToString());
            
            if (drs.Count() <= 0)
            {
                DataRow row = dt.NewRow();                
                row[dt.Columns[0].ColumnName] = dgv_Dept.SelectedRows[0].Cells[0].Value.ToString();
                dt.Rows.Add(row);
            }
            
            dgv_Tmpl_Dep_Ref.DataSource = dt;
        }
        public void RemoveRow(DataTable dt) {
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    if (dgv_Tmpl_Dep_Ref.SelectedRows[0] != null)
                    {
                        //dt.Rows.RemoveAt(dgv_Tmpl_Dep_Ref.SelectedRows[0].Index);
                        dt.Rows[dgv_Tmpl_Dep_Ref.SelectedRows[0].Index][1] = "V";
                    }
                }
                dgv_Tmpl_Dep_Ref.DataSource = dt;
            }
            catch (ArgumentOutOfRangeException ex)
            {
                MessageBox.Show("請選取資料列後再點選刪除");
            }
            catch (Exception ex) {
                MessageBox.Show("不明錯誤，請洽工程師");
            }
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            bool result = false;
            if (dgv_Tmpl_Dep_Ref.Rows.Count > 0)
            {
                ANBTK anbtk;
                List<ANBTK> list = new List<ANBTK>();

                for (int i = 0; i < dgv_Tmpl_Dep_Ref.Rows.Count; i++)
                {
                    anbtk = new ANBTK();
                    anbtk.Tk002 = dgv_Tmpl_Dep_Ref.Rows[i].Cells[0].Value.ToString();   // 部門代號
                    anbtk.Tk006 = dgv_Tmpl_Dep_Ref.Rows[i].Cells[1].Value.ToString();   // 刪除標記

                    if ("V".Equals(anbtk.Tk006)) // 若該筆已標記刪除
                    {
                        anbtk.Tk004 = "";   // 會科樣版編號，設為空值
                        anbtk.Tk005 = "";   // 會科樣版版本號，設為空值
                    }
                    else {
                        anbtk.Tk004 = cbx_Tmpl_ID.Text;   // 會科樣版編號
                        anbtk.Tk005 = cbx_Ver.Text;       // 會科樣版版本號
                    }

                    list.Add(anbtk);
                }

                result = Tmpl_Model.Update_Dept_Tmpl_Ref(list, cbx_Year.Text, gUserInfo);
            }
            else {
                MessageBox.Show("沒有資料可供更新");
                return;
            }

            if (result != true) MessageBox.Show("更新時出現問題，請洽工程師");
            else
            {
                // 重新載入
                string tmplID = cbx_Tmpl_ID.SelectedItem.ToString();
                string ver = cbx_Ver.SelectedItem.ToString();                
                Get_Tmpl_Dept_Ref_Table(tmplID, ver);   // 透過樣版編號，取得使用該樣版的部門列表
                
                MessageBox.Show("更新成功！");
            }

        }

        public void SetYear() {
            int year = DateTime.Today.Year;

            for (int i = year; i < year + 3; i++)
            {
                cbx_Year.Items.Add(i);
            }

            cbx_Year.SelectedIndex = 1;
        }
    }
}
