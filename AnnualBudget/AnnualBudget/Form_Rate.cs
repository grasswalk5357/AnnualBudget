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
    public partial class Form_Rate : Form
    {
        DataTable dt = null;
        public string gUserId = "";
        public string gDeptNo = "";

        public Form_Rate()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
        }

        private void Form_Rate_Load(object sender, EventArgs e)
        {
            SetRateContent();
        }

        /// <summary>
        /// 取得並設定財會維護的各項費率
        /// </summary>
        private void SetRateContent() {
            dt = ANBTJ_Model.GetRateTable();
            if (dt != null)
            {
                dgv_Rate.DataSource = dt;
                for (int i = 0; i < dgv_Rate.Columns.Count; i++)
                {
                    dgv_Rate.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    if (i <= 0) { 
                        dgv_Rate.Columns[i].ReadOnly = true;
                        dgv_Rate.Columns[i].DefaultCellStyle.BackColor = Color.PaleGreen;
                    }

                }
            }
        }

        private void btn_Edit_Click(object sender, EventArgs e)
        {
            dgv_Rate.Enabled = true;
            btn_Edit.Enabled = false;
            btn_AddRow.Enabled = true;
            btn_CancelEdit.Enabled = true;
            btn_Save.Enabled = true;
        }

        private void btn_CancelEdit_Click(object sender, EventArgs e)
        {
            SetBTNtoDefault();
            SetRateContent();
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {

            List<ANBTJ> list = new List<ANBTJ>();
            ANBTJ obj = null;
            bool result = false;
            
            for (int i = 0; i < dgv_Rate.Rows.Count; i++) 
            {
                if ("V".Equals(dgv_Rate.Rows[i].Cells[4].Value.ToString()))  // 只儲存有編輯過的資料行
                { 
                    obj = new ANBTJ();
                    obj.Tj002 = dgv_Rate.Rows[i].Cells[0].Value.ToString();         // 費率代號
                    obj.Tj003 = dgv_Rate.Rows[i].Cells[1].Value.ToString();         // 費率名稱
                    obj.Tj004 = Convert.ToDecimal(dgv_Rate.Rows[i].Cells[2].Value); // 費率
                    obj.Tj005 = dgv_Rate.Rows[i].Cells[3].Value.ToString();         // 備註

                    list.Add(obj);
                }
            }

            result = ANBTJ_Model.Update_Rate(list, gUserId);

            if (result == true)
            {
                SetBTNtoDefault();
                MessageBox.Show("儲存成功，請直接關閉視窗");
                
            }
            else
                MessageBox.Show("儲存失敗，請洽工程師");

            SetRateContent();
           
        }

        /// <summary>
        /// 將各按鍵的Enabled模式設為預設值
        /// </summary>
        private void SetBTNtoDefault() {
            
            dgv_Rate.Enabled = false;
            btn_Edit.Enabled = true;
            btn_AddRow.Enabled = false;
            btn_CancelEdit.Enabled = false;
            btn_Save.Enabled = false;
        }

        private void dgv_Rate_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dgv_Rate.Rows[e.RowIndex].Cells[4].Value = "V";
        }

        
        private void btn_AddRow_Click(object sender, EventArgs e)
        {
            string sNum = dt.Rows[dt.Rows.Count - 1][0].ToString().Substring(1);    // 取得目前費率代號的數字部份
            int iNum = Convert.ToInt32(sNum);   // 轉換成int
            string newItemNum = (iNum + 1).ToString();
            int len = newItemNum.Length;
            string newItemCode = "R";


            for (int i = 0; i < (4 - len); i++)
            {
                newItemNum = "0" + newItemNum;
            }
            newItemCode = newItemCode + newItemNum;

            dt.Rows.Add();

            dt.Rows[dt.Rows.Count - 1][0] = newItemCode;
            dgv_Rate.DataSource = dt;
        }
    }
}
