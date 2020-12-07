namespace AnnualBudget
{
    partial class Form_Rate
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_Rate));
            this.btn_Save = new System.Windows.Forms.Button();
            this.tbx_LaborIsr = new System.Windows.Forms.TextBox();
            this.lbl_LaborIsr = new System.Windows.Forms.Label();
            this.lbl_HealthIsr = new System.Windows.Forms.Label();
            this.tbx_HealthIsr = new System.Windows.Forms.TextBox();
            this.lbl_Food = new System.Windows.Forms.Label();
            this.tbx_Food = new System.Windows.Forms.TextBox();
            this.btn_CancelEdit = new System.Windows.Forms.Button();
            this.btn_Edit = new System.Windows.Forms.Button();
            this.dgv_Rate = new System.Windows.Forms.DataGridView();
            this.lbl_Notice = new System.Windows.Forms.Label();
            this.btn_AddRow = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Rate)).BeginInit();
            this.SuspendLayout();
            // 
            // btn_Save
            // 
            this.btn_Save.Enabled = false;
            this.btn_Save.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Save.Location = new System.Drawing.Point(405, 294);
            this.btn_Save.Name = "btn_Save";
            this.btn_Save.Size = new System.Drawing.Size(83, 27);
            this.btn_Save.TabIndex = 0;
            this.btn_Save.Text = "儲存";
            this.btn_Save.UseVisualStyleBackColor = true;
            this.btn_Save.Click += new System.EventHandler(this.btn_Save_Click);
            // 
            // tbx_LaborIsr
            // 
            this.tbx_LaborIsr.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tbx_LaborIsr.Location = new System.Drawing.Point(585, 229);
            this.tbx_LaborIsr.Name = "tbx_LaborIsr";
            this.tbx_LaborIsr.Size = new System.Drawing.Size(100, 27);
            this.tbx_LaborIsr.TabIndex = 1;
            // 
            // lbl_LaborIsr
            // 
            this.lbl_LaborIsr.AutoSize = true;
            this.lbl_LaborIsr.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_LaborIsr.Location = new System.Drawing.Point(510, 232);
            this.lbl_LaborIsr.Name = "lbl_LaborIsr";
            this.lbl_LaborIsr.Size = new System.Drawing.Size(69, 19);
            this.lbl_LaborIsr.TabIndex = 2;
            this.lbl_LaborIsr.Text = "勞保費率";
            // 
            // lbl_HealthIsr
            // 
            this.lbl_HealthIsr.AutoSize = true;
            this.lbl_HealthIsr.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_HealthIsr.Location = new System.Drawing.Point(510, 265);
            this.lbl_HealthIsr.Name = "lbl_HealthIsr";
            this.lbl_HealthIsr.Size = new System.Drawing.Size(69, 19);
            this.lbl_HealthIsr.TabIndex = 4;
            this.lbl_HealthIsr.Text = "健保費率";
            // 
            // tbx_HealthIsr
            // 
            this.tbx_HealthIsr.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tbx_HealthIsr.Location = new System.Drawing.Point(585, 262);
            this.tbx_HealthIsr.Name = "tbx_HealthIsr";
            this.tbx_HealthIsr.Size = new System.Drawing.Size(100, 27);
            this.tbx_HealthIsr.TabIndex = 3;
            // 
            // lbl_Food
            // 
            this.lbl_Food.AutoSize = true;
            this.lbl_Food.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_Food.Location = new System.Drawing.Point(510, 298);
            this.lbl_Food.Name = "lbl_Food";
            this.lbl_Food.Size = new System.Drawing.Size(69, 19);
            this.lbl_Food.TabIndex = 6;
            this.lbl_Food.Text = "伙食費率";
            // 
            // tbx_Food
            // 
            this.tbx_Food.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tbx_Food.Location = new System.Drawing.Point(585, 295);
            this.tbx_Food.Name = "tbx_Food";
            this.tbx_Food.Size = new System.Drawing.Size(100, 27);
            this.tbx_Food.TabIndex = 5;
            // 
            // btn_CancelEdit
            // 
            this.btn_CancelEdit.Enabled = false;
            this.btn_CancelEdit.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_CancelEdit.Location = new System.Drawing.Point(227, 294);
            this.btn_CancelEdit.Name = "btn_CancelEdit";
            this.btn_CancelEdit.Size = new System.Drawing.Size(83, 27);
            this.btn_CancelEdit.TabIndex = 7;
            this.btn_CancelEdit.Text = "取消編輯";
            this.btn_CancelEdit.UseVisualStyleBackColor = true;
            this.btn_CancelEdit.Click += new System.EventHandler(this.btn_CancelEdit_Click);
            // 
            // btn_Edit
            // 
            this.btn_Edit.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Edit.Location = new System.Drawing.Point(138, 294);
            this.btn_Edit.Name = "btn_Edit";
            this.btn_Edit.Size = new System.Drawing.Size(83, 27);
            this.btn_Edit.TabIndex = 8;
            this.btn_Edit.Text = "編輯";
            this.btn_Edit.UseVisualStyleBackColor = true;
            this.btn_Edit.Click += new System.EventHandler(this.btn_Edit_Click);
            // 
            // dgv_Rate
            // 
            this.dgv_Rate.AllowUserToAddRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_Rate.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv_Rate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_Rate.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgv_Rate.Enabled = false;
            this.dgv_Rate.Location = new System.Drawing.Point(29, 22);
            this.dgv_Rate.Name = "dgv_Rate";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_Rate.RowHeadersDefaultCellStyle = dataGridViewCellStyle2;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.dgv_Rate.RowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dgv_Rate.RowTemplate.Height = 24;
            this.dgv_Rate.Size = new System.Drawing.Size(459, 249);
            this.dgv_Rate.TabIndex = 9;
            this.dgv_Rate.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_Rate_CellEndEdit);
            // 
            // lbl_Notice
            // 
            this.lbl_Notice.AutoSize = true;
            this.lbl_Notice.Font = new System.Drawing.Font("微軟正黑體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_Notice.Location = new System.Drawing.Point(27, 274);
            this.lbl_Notice.Name = "lbl_Notice";
            this.lbl_Notice.Size = new System.Drawing.Size(124, 16);
            this.lbl_Notice.TabIndex = 10;
            this.lbl_Notice.Text = "* 僅白色區域可供編輯";
            // 
            // btn_AddRow
            // 
            this.btn_AddRow.Enabled = false;
            this.btn_AddRow.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_AddRow.Location = new System.Drawing.Point(316, 294);
            this.btn_AddRow.Name = "btn_AddRow";
            this.btn_AddRow.Size = new System.Drawing.Size(83, 27);
            this.btn_AddRow.TabIndex = 11;
            this.btn_AddRow.Text = "新增一筆";
            this.btn_AddRow.UseVisualStyleBackColor = true;
            this.btn_AddRow.Click += new System.EventHandler(this.btn_AddRow_Click);
            // 
            // Form_Rate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(513, 356);
            this.Controls.Add(this.btn_AddRow);
            this.Controls.Add(this.lbl_Notice);
            this.Controls.Add(this.dgv_Rate);
            this.Controls.Add(this.btn_Edit);
            this.Controls.Add(this.btn_CancelEdit);
            this.Controls.Add(this.lbl_Food);
            this.Controls.Add(this.tbx_Food);
            this.Controls.Add(this.lbl_HealthIsr);
            this.Controls.Add(this.tbx_HealthIsr);
            this.Controls.Add(this.lbl_LaborIsr);
            this.Controls.Add(this.tbx_LaborIsr);
            this.Controls.Add(this.btn_Save);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form_Rate";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "各項費率調整";
            this.Load += new System.EventHandler(this.Form_Rate_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Rate)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Save;
        private System.Windows.Forms.TextBox tbx_LaborIsr;
        private System.Windows.Forms.Label lbl_LaborIsr;
        private System.Windows.Forms.Label lbl_HealthIsr;
        private System.Windows.Forms.TextBox tbx_HealthIsr;
        private System.Windows.Forms.Label lbl_Food;
        private System.Windows.Forms.TextBox tbx_Food;
        private System.Windows.Forms.Button btn_CancelEdit;
        private System.Windows.Forms.Button btn_Edit;
        private System.Windows.Forms.DataGridView dgv_Rate;
        private System.Windows.Forms.Label lbl_Notice;
        private System.Windows.Forms.Button btn_AddRow;
    }
}