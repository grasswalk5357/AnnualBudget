namespace AnnualBudget
{
    partial class Form_Login
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_Login));
            this.lbl_userName = new System.Windows.Forms.Label();
            this.tbx_userName = new System.Windows.Forms.TextBox();
            this.tbx_PWD = new System.Windows.Forms.TextBox();
            this.lbl_PWD = new System.Windows.Forms.Label();
            this.btn_Login = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbl_userName
            // 
            this.lbl_userName.AutoSize = true;
            this.lbl_userName.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_userName.Location = new System.Drawing.Point(19, 69);
            this.lbl_userName.Name = "lbl_userName";
            this.lbl_userName.Size = new System.Drawing.Size(84, 19);
            this.lbl_userName.TabIndex = 0;
            this.lbl_userName.Text = "AD帳號：";
            // 
            // tbx_userName
            // 
            this.tbx_userName.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tbx_userName.Location = new System.Drawing.Point(102, 59);
            this.tbx_userName.Name = "tbx_userName";
            this.tbx_userName.Size = new System.Drawing.Size(147, 27);
            this.tbx_userName.TabIndex = 1;
            // 
            // tbx_PWD
            // 
            this.tbx_PWD.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tbx_PWD.Location = new System.Drawing.Point(102, 108);
            this.tbx_PWD.Name = "tbx_PWD";
            this.tbx_PWD.PasswordChar = '*';
            this.tbx_PWD.Size = new System.Drawing.Size(147, 27);
            this.tbx_PWD.TabIndex = 3;
            // 
            // lbl_PWD
            // 
            this.lbl_PWD.AutoSize = true;
            this.lbl_PWD.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_PWD.Location = new System.Drawing.Point(42, 111);
            this.lbl_PWD.Name = "lbl_PWD";
            this.lbl_PWD.Size = new System.Drawing.Size(54, 19);
            this.lbl_PWD.TabIndex = 2;
            this.lbl_PWD.Text = "密碼：";
            // 
            // btn_Login
            // 
            this.btn_Login.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Login.Location = new System.Drawing.Point(174, 165);
            this.btn_Login.Name = "btn_Login";
            this.btn_Login.Size = new System.Drawing.Size(75, 28);
            this.btn_Login.TabIndex = 4;
            this.btn_Login.Text = "登 入";
            this.btn_Login.UseVisualStyleBackColor = true;
            this.btn_Login.Click += new System.EventHandler(this.btn_Login_Click);
            // 
            // Form_Login
            // 
            this.AcceptButton = this.btn_Login;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(294, 231);
            this.Controls.Add(this.btn_Login);
            this.Controls.Add(this.tbx_PWD);
            this.Controls.Add(this.lbl_PWD);
            this.Controls.Add(this.tbx_userName);
            this.Controls.Add(this.lbl_userName);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form_Login";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "年度預算編列-使用者登入";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form_Login_FormClosed);
            this.Load += new System.EventHandler(this.Form_Login_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_userName;
        private System.Windows.Forms.TextBox tbx_userName;
        private System.Windows.Forms.TextBox tbx_PWD;
        private System.Windows.Forms.Label lbl_PWD;
        private System.Windows.Forms.Button btn_Login;
    }
}