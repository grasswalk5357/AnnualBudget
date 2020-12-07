using AnnualBudget.BOs;
using AnnualBudget.Model;
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
    public partial class Form_AuthMgmt : Form
    {
        private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public UserInfo gUserInfo = new UserInfo();

        public Form_AuthMgmt()
        {
            InitializeComponent();
        }

        private void Form_AuthMgmt_Load(object sender, EventArgs e)
        {
            _log.Debug("「人員權限管理」表單載入中");

            List<Object> list = ANBTO_Model.LoadData();
            SetData(list);
            SetBannedRowColor();

            _log.Debug("「人員權限管理」表單載入完畢");
        }

        private void SetData(List<Object> list)
        {
            dgv_AuthMgmt.Rows.Clear();

            if (list != null && list.Count > 0)
            {
                for (int i = 0; i < list.Count; i++)
                {
                    dgv_AuthMgmt.Rows.Add();

                    ANBTO anbto = (ANBTO)list[i];
                    dgv_AuthMgmt.Rows[i].Cells["dgv_AuthMgmt_TO001"].Value = anbto.To001;
                    dgv_AuthMgmt.Rows[i].Cells["dgv_AuthMgmt_TO002"].Value = anbto.To002;
                    dgv_AuthMgmt.Rows[i].Cells["dgv_AuthMgmt_TO003"].Value = anbto.To003;
                    dgv_AuthMgmt.Rows[i].Cells["dgv_AuthMgmt_TO004"].Value = anbto.To004;
                    dgv_AuthMgmt.Rows[i].Cells["dgv_AuthMgmt_TO005"].Value = anbto.To005;
                }
            }

        }

        private void btn_Edit_Click(object sender, EventArgs e)
        {
            GoEdit();
        }

        private void btn_Add_Click(object sender, EventArgs e)
        {
            bool checkResult = CheckInfo();
            bool checkDuplicate = false;

            if (dgv_AuthMgmt.RowCount > 0)
                checkDuplicate = CheckDuplicate();

            if (checkResult == true && checkDuplicate == true) {
                dgv_AuthMgmt.Rows.Add();
                dgv_AuthMgmt.Rows[dgv_AuthMgmt.RowCount - 1].Cells["dgv_AuthMgmt_TO002"].Value = tbx_EmpID.Text;
                dgv_AuthMgmt.Rows[dgv_AuthMgmt.RowCount - 1].Cells["dgv_AuthMgmt_TO003"].Value = tbx_EmpName.Text;
                dgv_AuthMgmt.Rows[dgv_AuthMgmt.RowCount - 1].Cells["dgv_AuthMgmt_TO004"].Value = tbx_EmpAD_ID.Text;
                dgv_AuthMgmt.Rows[dgv_AuthMgmt.RowCount - 1].Cells["dgv_AuthMgmt_TO005"].Value = "N";
                dgv_AuthMgmt.Rows[dgv_AuthMgmt.RowCount - 1].Cells["dgv_AuthMgmt_IsEdited"].Value = "V";
            }


        }

        private bool CheckInfo() {
            bool result = true;
            if (String.IsNullOrEmpty(tbx_EmpID.Text))
            {
                MessageBox.Show("「工號」不允許空白");
                result = false;
            }
            else if (String.IsNullOrEmpty(tbx_EmpAD_ID.Text))
            {
                MessageBox.Show("「AD帳號」不允許空白");
                result = false;
            }
            else if (String.IsNullOrEmpty(tbx_EmpName.Text))
            {
                MessageBox.Show("「姓名」不允許空白");
                result = false;
            }
            return result;
        }

        private bool CheckDuplicate() {
            bool result = true;
            string EmpId = tbx_EmpID.Text;
            string EmpAD_Id = tbx_EmpAD_ID.Text;

            for (int i = 0; i < dgv_AuthMgmt.RowCount; i++) {
                if (dgv_AuthMgmt.Rows[i].Cells["dgv_AuthMgmt_TO002"].Value != null) {
                    if (EmpId.Equals(dgv_AuthMgmt.Rows[i].Cells["dgv_AuthMgmt_TO002"].Value.ToString()))
                    {
                        result = false;
                        MessageBox.Show("工號：" + tbx_EmpID.Text + "已存在，請重新檢查");
                        break;
                    }

                }
                
                if (dgv_AuthMgmt.Rows[i].Cells["dgv_AuthMgmt_TO004"].Value != null)
                {
                    if (EmpAD_Id.Equals(dgv_AuthMgmt.Rows[i].Cells["dgv_AuthMgmt_TO004"].Value.ToString())) {
                        result = false;
                        MessageBox.Show("AD帳號：" + tbx_EmpAD_ID.Text + "已存在，請重新檢查");
                        break;
                    }
                }
            }

            return result;
        }

        private void btn_CancelEdit_Click(object sender, EventArgs e)
        {
            List<Object> list = ANBTO_Model.LoadData();
            SetData(list);

            CancelEdit();            
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            List<Object> list = ANBTO_Model.SaveMatrixData(dgv_AuthMgmt, gUserInfo.UserID, gUserInfo.DeptNo);
            bool result = ANBTO_Model.UpdateToANBTO(list);

            if (result == true)
            {
                MessageBox.Show("資料儲存成功");
                List<Object> mylist = ANBTO_Model.LoadData();
                SetData(mylist);
                SetBannedRowColor();
                CancelEdit();
                
            }
            else
            {
                MessageBox.Show("資料儲存失敗！請重新檢查後再次儲存");
            }
        }

        private void CancelEdit() {
            dgv_AuthMgmt.Enabled = false;
            btn_Edit.Enabled = true;
            btn_Add.Enabled = false;
            btn_CancelEdit.Enabled = false;
            btn_Save.Enabled = false;

            tbx_EmpAD_ID.Enabled = false;
            tbx_EmpAD_ID.Text = "";
            tbx_EmpID.Enabled = false;
            tbx_EmpID.Text = "";

            tbx_EmpName.Enabled = false;
            tbx_EmpName.Text = "";
        }

        private void GoEdit() {
            dgv_AuthMgmt.Enabled = true;
            btn_Edit.Enabled = false;
            btn_Add.Enabled = true;
            btn_CancelEdit.Enabled = true;
            btn_Save.Enabled = true;

            tbx_EmpAD_ID.Enabled = true;
            tbx_EmpID.Enabled = true;
            tbx_EmpName.Enabled = true;
        }

        private void dgv_AuthMgmt_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dgv_AuthMgmt.Rows[e.RowIndex].Cells["dgv_AuthMgmt_IsEdited"].Value = "V";
        }

        private void SetBannedRowColor() {
            if (dgv_AuthMgmt.RowCount > 0) {
                for (int i = 0; i < dgv_AuthMgmt.RowCount; i++) {
                    if (dgv_AuthMgmt.Rows[i].Cells["dgv_AuthMgmt_TO005"].Value != null) {
                        if ("Y".Equals(dgv_AuthMgmt.Rows[i].Cells["dgv_AuthMgmt_TO005"].Value.ToString())) {
                            dgv_AuthMgmt.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;
                        }
                    }
                }
            }
        }
    }
}
