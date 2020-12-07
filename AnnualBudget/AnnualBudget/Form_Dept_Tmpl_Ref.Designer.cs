namespace AnnualBudget
{
    partial class Form_Dept_Tmpl_Ref
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_Dept_Tmpl_Ref));
            this.cbx_Tmpl_ID = new System.Windows.Forms.ComboBox();
            this.cbx_Ver = new System.Windows.Forms.ComboBox();
            this.btn_SearchTmpl = new System.Windows.Forms.Button();
            this.dgv_Tmpl = new System.Windows.Forms.DataGridView();
            this.dgv_Dept = new System.Windows.Forms.DataGridView();
            this.dgv_Tmpl_Dep_Ref = new System.Windows.Forms.DataGridView();
            this.lbl_Tmpl_ID = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lbl_Dept = new System.Windows.Forms.Label();
            this.lbl_Tmpl_Dep_Ref = new System.Windows.Forms.Label();
            this.btn_Add = new System.Windows.Forms.Button();
            this.btn_Remove = new System.Windows.Forms.Button();
            this.btn_Save = new System.Windows.Forms.Button();
            this.lbl_TmplName = new System.Windows.Forms.Label();
            this.tbx_TmplName = new System.Windows.Forms.TextBox();
            this.lbl_TmplContent = new System.Windows.Forms.Label();
            this.lbl_Year = new System.Windows.Forms.Label();
            this.cbx_Year = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Tmpl)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Dept)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Tmpl_Dep_Ref)).BeginInit();
            this.SuspendLayout();
            // 
            // cbx_Tmpl_ID
            // 
            this.cbx_Tmpl_ID.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbx_Tmpl_ID.FormattingEnabled = true;
            this.cbx_Tmpl_ID.Location = new System.Drawing.Point(106, 27);
            this.cbx_Tmpl_ID.Name = "cbx_Tmpl_ID";
            this.cbx_Tmpl_ID.Size = new System.Drawing.Size(100, 27);
            this.cbx_Tmpl_ID.TabIndex = 0;
            this.cbx_Tmpl_ID.SelectedIndexChanged += new System.EventHandler(this.cbx_Tmpl_ID_SelectedIndexChanged);
            // 
            // cbx_Ver
            // 
            this.cbx_Ver.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbx_Ver.FormattingEnabled = true;
            this.cbx_Ver.Location = new System.Drawing.Point(106, 60);
            this.cbx_Ver.Name = "cbx_Ver";
            this.cbx_Ver.Size = new System.Drawing.Size(100, 27);
            this.cbx_Ver.TabIndex = 1;
            // 
            // btn_SearchTmpl
            // 
            this.btn_SearchTmpl.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_SearchTmpl.Location = new System.Drawing.Point(347, 93);
            this.btn_SearchTmpl.Name = "btn_SearchTmpl";
            this.btn_SearchTmpl.Size = new System.Drawing.Size(82, 26);
            this.btn_SearchTmpl.TabIndex = 2;
            this.btn_SearchTmpl.Text = "搜尋模版";
            this.btn_SearchTmpl.UseVisualStyleBackColor = true;
            this.btn_SearchTmpl.Click += new System.EventHandler(this.btn_SearchTmpl_Click);
            // 
            // dgv_Tmpl
            // 
            this.dgv_Tmpl.AllowUserToAddRows = false;
            this.dgv_Tmpl.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_Tmpl.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv_Tmpl.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv_Tmpl.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgv_Tmpl.Location = new System.Drawing.Point(21, 125);
            this.dgv_Tmpl.Name = "dgv_Tmpl";
            this.dgv_Tmpl.ReadOnly = true;
            this.dgv_Tmpl.RowTemplate.Height = 24;
            this.dgv_Tmpl.Size = new System.Drawing.Size(408, 294);
            this.dgv_Tmpl.TabIndex = 3;
            // 
            // dgv_Dept
            // 
            this.dgv_Dept.AllowUserToAddRows = false;
            this.dgv_Dept.AllowUserToDeleteRows = false;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_Dept.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dgv_Dept.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv_Dept.DefaultCellStyle = dataGridViewCellStyle4;
            this.dgv_Dept.Location = new System.Drawing.Point(506, 60);
            this.dgv_Dept.Name = "dgv_Dept";
            this.dgv_Dept.ReadOnly = true;
            this.dgv_Dept.RowTemplate.Height = 24;
            this.dgv_Dept.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_Dept.Size = new System.Drawing.Size(260, 359);
            this.dgv_Dept.TabIndex = 6;
            // 
            // dgv_Tmpl_Dep_Ref
            // 
            this.dgv_Tmpl_Dep_Ref.AllowUserToAddRows = false;
            this.dgv_Tmpl_Dep_Ref.AllowUserToDeleteRows = false;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_Tmpl_Dep_Ref.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle5;
            this.dgv_Tmpl_Dep_Ref.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv_Tmpl_Dep_Ref.DefaultCellStyle = dataGridViewCellStyle6;
            this.dgv_Tmpl_Dep_Ref.Location = new System.Drawing.Point(826, 60);
            this.dgv_Tmpl_Dep_Ref.Name = "dgv_Tmpl_Dep_Ref";
            this.dgv_Tmpl_Dep_Ref.ReadOnly = true;
            this.dgv_Tmpl_Dep_Ref.RowTemplate.Height = 24;
            this.dgv_Tmpl_Dep_Ref.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_Tmpl_Dep_Ref.Size = new System.Drawing.Size(240, 357);
            this.dgv_Tmpl_Dep_Ref.TabIndex = 7;
            // 
            // lbl_Tmpl_ID
            // 
            this.lbl_Tmpl_ID.AutoSize = true;
            this.lbl_Tmpl_ID.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_Tmpl_ID.Location = new System.Drawing.Point(32, 30);
            this.lbl_Tmpl_ID.Name = "lbl_Tmpl_ID";
            this.lbl_Tmpl_ID.Size = new System.Drawing.Size(69, 19);
            this.lbl_Tmpl_ID.TabIndex = 8;
            this.lbl_Tmpl_ID.Text = "模版編號";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(17, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 19);
            this.label1.TabIndex = 9;
            this.label1.Text = "模版版本號";
            // 
            // lbl_Dept
            // 
            this.lbl_Dept.AutoSize = true;
            this.lbl_Dept.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_Dept.Location = new System.Drawing.Point(502, 35);
            this.lbl_Dept.Name = "lbl_Dept";
            this.lbl_Dept.Size = new System.Drawing.Size(99, 19);
            this.lbl_Dept.TabIndex = 10;
            this.lbl_Dept.Text = "公司部門列表";
            // 
            // lbl_Tmpl_Dep_Ref
            // 
            this.lbl_Tmpl_Dep_Ref.AutoSize = true;
            this.lbl_Tmpl_Dep_Ref.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_Tmpl_Dep_Ref.Location = new System.Drawing.Point(822, 35);
            this.lbl_Tmpl_Dep_Ref.Name = "lbl_Tmpl_Dep_Ref";
            this.lbl_Tmpl_Dep_Ref.Size = new System.Drawing.Size(144, 19);
            this.lbl_Tmpl_Dep_Ref.TabIndex = 11;
            this.lbl_Tmpl_Dep_Ref.Text = "套用模版之部門列表";
            // 
            // btn_Add
            // 
            this.btn_Add.Enabled = false;
            this.btn_Add.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Add.Location = new System.Drawing.Point(773, 176);
            this.btn_Add.Name = "btn_Add";
            this.btn_Add.Size = new System.Drawing.Size(47, 29);
            this.btn_Add.TabIndex = 12;
            this.btn_Add.Text = "＋";
            this.btn_Add.UseVisualStyleBackColor = true;
            this.btn_Add.Click += new System.EventHandler(this.btn_Add_Click);
            // 
            // btn_Remove
            // 
            this.btn_Remove.Enabled = false;
            this.btn_Remove.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Remove.Location = new System.Drawing.Point(773, 211);
            this.btn_Remove.Name = "btn_Remove";
            this.btn_Remove.Size = new System.Drawing.Size(47, 29);
            this.btn_Remove.TabIndex = 13;
            this.btn_Remove.Text = "－";
            this.btn_Remove.UseVisualStyleBackColor = true;
            this.btn_Remove.Click += new System.EventHandler(this.btn_Remove_Click);
            // 
            // btn_Save
            // 
            this.btn_Save.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Save.Location = new System.Drawing.Point(889, 434);
            this.btn_Save.Name = "btn_Save";
            this.btn_Save.Size = new System.Drawing.Size(82, 26);
            this.btn_Save.TabIndex = 15;
            this.btn_Save.Text = "存　檔";
            this.btn_Save.UseVisualStyleBackColor = true;
            this.btn_Save.Click += new System.EventHandler(this.btn_Save_Click);
            // 
            // lbl_TmplName
            // 
            this.lbl_TmplName.AutoSize = true;
            this.lbl_TmplName.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_TmplName.Location = new System.Drawing.Point(254, 30);
            this.lbl_TmplName.Name = "lbl_TmplName";
            this.lbl_TmplName.Size = new System.Drawing.Size(69, 19);
            this.lbl_TmplName.TabIndex = 17;
            this.lbl_TmplName.Text = "模版名稱";
            // 
            // tbx_TmplName
            // 
            this.tbx_TmplName.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tbx_TmplName.Location = new System.Drawing.Point(329, 27);
            this.tbx_TmplName.Name = "tbx_TmplName";
            this.tbx_TmplName.Size = new System.Drawing.Size(100, 27);
            this.tbx_TmplName.TabIndex = 18;
            // 
            // lbl_TmplContent
            // 
            this.lbl_TmplContent.AutoSize = true;
            this.lbl_TmplContent.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_TmplContent.Location = new System.Drawing.Point(17, 101);
            this.lbl_TmplContent.Name = "lbl_TmplContent";
            this.lbl_TmplContent.Size = new System.Drawing.Size(159, 19);
            this.lbl_TmplContent.TabIndex = 19;
            this.lbl_TmplContent.Text = "預算總表模版內容列表";
            // 
            // lbl_Year
            // 
            this.lbl_Year.AutoSize = true;
            this.lbl_Year.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_Year.Location = new System.Drawing.Point(284, 63);
            this.lbl_Year.Name = "lbl_Year";
            this.lbl_Year.Size = new System.Drawing.Size(39, 19);
            this.lbl_Year.TabIndex = 21;
            this.lbl_Year.Text = "年度";
            // 
            // cbx_Year
            // 
            this.cbx_Year.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbx_Year.FormattingEnabled = true;
            this.cbx_Year.Location = new System.Drawing.Point(329, 60);
            this.cbx_Year.Name = "cbx_Year";
            this.cbx_Year.Size = new System.Drawing.Size(100, 27);
            this.cbx_Year.TabIndex = 20;
            // 
            // Form_Dept_Tmpl_Ref
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1108, 490);
            this.Controls.Add(this.lbl_Year);
            this.Controls.Add(this.cbx_Year);
            this.Controls.Add(this.lbl_TmplContent);
            this.Controls.Add(this.tbx_TmplName);
            this.Controls.Add(this.lbl_TmplName);
            this.Controls.Add(this.btn_Save);
            this.Controls.Add(this.btn_Remove);
            this.Controls.Add(this.btn_Add);
            this.Controls.Add(this.lbl_Tmpl_Dep_Ref);
            this.Controls.Add(this.lbl_Dept);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lbl_Tmpl_ID);
            this.Controls.Add(this.dgv_Tmpl_Dep_Ref);
            this.Controls.Add(this.dgv_Dept);
            this.Controls.Add(this.dgv_Tmpl);
            this.Controls.Add(this.btn_SearchTmpl);
            this.Controls.Add(this.cbx_Ver);
            this.Controls.Add(this.cbx_Tmpl_ID);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form_Dept_Tmpl_Ref";
            this.Text = "會科預算總表樣版與部門對應設計";
            this.Load += new System.EventHandler(this.Form_Dept_Tmpl_Ref_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Tmpl)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Dept)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Tmpl_Dep_Ref)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbx_Tmpl_ID;
        private System.Windows.Forms.ComboBox cbx_Ver;
        private System.Windows.Forms.Button btn_SearchTmpl;
        private System.Windows.Forms.DataGridView dgv_Tmpl;
        private System.Windows.Forms.DataGridView dgv_Dept;
        private System.Windows.Forms.DataGridView dgv_Tmpl_Dep_Ref;
        private System.Windows.Forms.Label lbl_Tmpl_ID;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lbl_Dept;
        private System.Windows.Forms.Label lbl_Tmpl_Dep_Ref;
        private System.Windows.Forms.Button btn_Add;
        private System.Windows.Forms.Button btn_Remove;
        private System.Windows.Forms.Button btn_Save;
        private System.Windows.Forms.Label lbl_TmplName;
        private System.Windows.Forms.TextBox tbx_TmplName;
        private System.Windows.Forms.Label lbl_TmplContent;
        private System.Windows.Forms.Label lbl_Year;
        private System.Windows.Forms.ComboBox cbx_Year;
    }
}