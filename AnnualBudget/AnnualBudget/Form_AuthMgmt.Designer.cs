namespace AnnualBudget
{
    partial class Form_AuthMgmt
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_AuthMgmt));
            this.btn_Add = new System.Windows.Forms.Button();
            this.btn_Edit = new System.Windows.Forms.Button();
            this.btn_CancelEdit = new System.Windows.Forms.Button();
            this.btn_Save = new System.Windows.Forms.Button();
            this.dgv_AuthMgmt = new System.Windows.Forms.DataGridView();
            this.dgv_AuthMgmt_TO001 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgv_AuthMgmt_TO002 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgv_AuthMgmt_TO003 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgv_AuthMgmt_TO004 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgv_AuthMgmt_TO005 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgv_AuthMgmt_IsEdited = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tbx_EmpID = new System.Windows.Forms.TextBox();
            this.tbx_EmpAD_ID = new System.Windows.Forms.TextBox();
            this.tbx_EmpName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_AuthMgmt)).BeginInit();
            this.SuspendLayout();
            // 
            // btn_Add
            // 
            this.btn_Add.Enabled = false;
            this.btn_Add.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Add.Location = new System.Drawing.Point(207, 438);
            this.btn_Add.Name = "btn_Add";
            this.btn_Add.Size = new System.Drawing.Size(111, 27);
            this.btn_Add.TabIndex = 15;
            this.btn_Add.Text = "新增人員權限";
            this.btn_Add.UseVisualStyleBackColor = true;
            this.btn_Add.Click += new System.EventHandler(this.btn_Add_Click);
            // 
            // btn_Edit
            // 
            this.btn_Edit.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Edit.Location = new System.Drawing.Point(391, 406);
            this.btn_Edit.Name = "btn_Edit";
            this.btn_Edit.Size = new System.Drawing.Size(83, 27);
            this.btn_Edit.TabIndex = 14;
            this.btn_Edit.Text = "啟用編輯";
            this.btn_Edit.UseVisualStyleBackColor = true;
            this.btn_Edit.Click += new System.EventHandler(this.btn_Edit_Click);
            // 
            // btn_CancelEdit
            // 
            this.btn_CancelEdit.Enabled = false;
            this.btn_CancelEdit.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_CancelEdit.Location = new System.Drawing.Point(391, 438);
            this.btn_CancelEdit.Name = "btn_CancelEdit";
            this.btn_CancelEdit.Size = new System.Drawing.Size(83, 27);
            this.btn_CancelEdit.TabIndex = 13;
            this.btn_CancelEdit.Text = "取消編輯";
            this.btn_CancelEdit.UseVisualStyleBackColor = true;
            this.btn_CancelEdit.Click += new System.EventHandler(this.btn_CancelEdit_Click);
            // 
            // btn_Save
            // 
            this.btn_Save.Enabled = false;
            this.btn_Save.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Save.Location = new System.Drawing.Point(494, 438);
            this.btn_Save.Name = "btn_Save";
            this.btn_Save.Size = new System.Drawing.Size(83, 27);
            this.btn_Save.TabIndex = 12;
            this.btn_Save.Text = "儲存設定";
            this.btn_Save.UseVisualStyleBackColor = true;
            this.btn_Save.Click += new System.EventHandler(this.btn_Save_Click);
            // 
            // dgv_AuthMgmt
            // 
            this.dgv_AuthMgmt.AllowUserToAddRows = false;
            this.dgv_AuthMgmt.AllowUserToDeleteRows = false;
            this.dgv_AuthMgmt.AllowUserToResizeColumns = false;
            this.dgv_AuthMgmt.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_AuthMgmt.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv_AuthMgmt.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_AuthMgmt.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dgv_AuthMgmt_TO001,
            this.dgv_AuthMgmt_TO002,
            this.dgv_AuthMgmt_TO003,
            this.dgv_AuthMgmt_TO004,
            this.dgv_AuthMgmt_TO005,
            this.dgv_AuthMgmt_IsEdited});
            this.dgv_AuthMgmt.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgv_AuthMgmt.Enabled = false;
            this.dgv_AuthMgmt.Location = new System.Drawing.Point(12, 12);
            this.dgv_AuthMgmt.Name = "dgv_AuthMgmt";
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_AuthMgmt.RowHeadersDefaultCellStyle = dataGridViewCellStyle7;
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle8.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.dgv_AuthMgmt.RowsDefaultCellStyle = dataGridViewCellStyle8;
            this.dgv_AuthMgmt.RowTemplate.Height = 24;
            this.dgv_AuthMgmt.Size = new System.Drawing.Size(565, 351);
            this.dgv_AuthMgmt.TabIndex = 16;
            this.dgv_AuthMgmt.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_AuthMgmt_CellEndEdit);
            // 
            // dgv_AuthMgmt_TO001
            // 
            this.dgv_AuthMgmt_TO001.HeaderText = "Index";
            this.dgv_AuthMgmt_TO001.Name = "dgv_AuthMgmt_TO001";
            this.dgv_AuthMgmt_TO001.ReadOnly = true;
            this.dgv_AuthMgmt_TO001.Visible = false;
            // 
            // dgv_AuthMgmt_TO002
            // 
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgv_AuthMgmt_TO002.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgv_AuthMgmt_TO002.HeaderText = "工號";
            this.dgv_AuthMgmt_TO002.Name = "dgv_AuthMgmt_TO002";
            this.dgv_AuthMgmt_TO002.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dgv_AuthMgmt_TO003
            // 
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.dgv_AuthMgmt_TO003.DefaultCellStyle = dataGridViewCellStyle3;
            this.dgv_AuthMgmt_TO003.HeaderText = "姓名";
            this.dgv_AuthMgmt_TO003.Name = "dgv_AuthMgmt_TO003";
            this.dgv_AuthMgmt_TO003.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dgv_AuthMgmt_TO004
            // 
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgv_AuthMgmt_TO004.DefaultCellStyle = dataGridViewCellStyle4;
            this.dgv_AuthMgmt_TO004.HeaderText = "AD帳號";
            this.dgv_AuthMgmt_TO004.Name = "dgv_AuthMgmt_TO004";
            this.dgv_AuthMgmt_TO004.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dgv_AuthMgmt_TO005
            // 
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.dgv_AuthMgmt_TO005.DefaultCellStyle = dataGridViewCellStyle5;
            this.dgv_AuthMgmt_TO005.HeaderText = "權限是否禁止";
            this.dgv_AuthMgmt_TO005.Name = "dgv_AuthMgmt_TO005";
            this.dgv_AuthMgmt_TO005.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dgv_AuthMgmt_IsEdited
            // 
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.dgv_AuthMgmt_IsEdited.DefaultCellStyle = dataGridViewCellStyle6;
            this.dgv_AuthMgmt_IsEdited.HeaderText = "已編輯";
            this.dgv_AuthMgmt_IsEdited.Name = "dgv_AuthMgmt_IsEdited";
            this.dgv_AuthMgmt_IsEdited.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // tbx_EmpID
            // 
            this.tbx_EmpID.Enabled = false;
            this.tbx_EmpID.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tbx_EmpID.Location = new System.Drawing.Point(88, 378);
            this.tbx_EmpID.Name = "tbx_EmpID";
            this.tbx_EmpID.Size = new System.Drawing.Size(100, 25);
            this.tbx_EmpID.TabIndex = 17;
            // 
            // tbx_EmpAD_ID
            // 
            this.tbx_EmpAD_ID.Enabled = false;
            this.tbx_EmpAD_ID.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tbx_EmpAD_ID.Location = new System.Drawing.Point(88, 409);
            this.tbx_EmpAD_ID.Name = "tbx_EmpAD_ID";
            this.tbx_EmpAD_ID.Size = new System.Drawing.Size(100, 25);
            this.tbx_EmpAD_ID.TabIndex = 18;
            // 
            // tbx_EmpName
            // 
            this.tbx_EmpName.Enabled = false;
            this.tbx_EmpName.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tbx_EmpName.Location = new System.Drawing.Point(88, 440);
            this.tbx_EmpName.Name = "tbx_EmpName";
            this.tbx_EmpName.Size = new System.Drawing.Size(100, 25);
            this.tbx_EmpName.TabIndex = 19;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(28, 379);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 19);
            this.label1.TabIndex = 20;
            this.label1.Text = "工號：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(7, 410);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 19);
            this.label2.TabIndex = 21;
            this.label2.Text = "AD帳號：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(28, 441);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 19);
            this.label3.TabIndex = 22;
            this.label3.Text = "姓名：";
            // 
            // Form_AuthMgmt
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(596, 485);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbx_EmpName);
            this.Controls.Add(this.tbx_EmpAD_ID);
            this.Controls.Add(this.tbx_EmpID);
            this.Controls.Add(this.dgv_AuthMgmt);
            this.Controls.Add(this.btn_Add);
            this.Controls.Add(this.btn_Edit);
            this.Controls.Add(this.btn_CancelEdit);
            this.Controls.Add(this.btn_Save);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form_AuthMgmt";
            this.Text = "人員權限管理";
            this.Load += new System.EventHandler(this.Form_AuthMgmt_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_AuthMgmt)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Add;
        private System.Windows.Forms.Button btn_Edit;
        private System.Windows.Forms.Button btn_CancelEdit;
        private System.Windows.Forms.Button btn_Save;
        private System.Windows.Forms.DataGridView dgv_AuthMgmt;
        private System.Windows.Forms.TextBox tbx_EmpID;
        private System.Windows.Forms.TextBox tbx_EmpAD_ID;
        private System.Windows.Forms.TextBox tbx_EmpName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgv_AuthMgmt_TO001;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgv_AuthMgmt_TO002;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgv_AuthMgmt_TO003;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgv_AuthMgmt_TO004;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgv_AuthMgmt_TO005;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgv_AuthMgmt_IsEdited;
    }
}