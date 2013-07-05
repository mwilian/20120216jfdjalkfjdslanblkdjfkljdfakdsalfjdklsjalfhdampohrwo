namespace dCube
{
    partial class FrmConnection
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmConnection));
            this.btnOKQD = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.User = new System.Windows.Forms.TextBox();
            this.Pass = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Server = new System.Windows.Forms.ComboBox();
            this.Database = new System.Windows.Forms.ComboBox();
            this.txtGeneralTimeOut = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.txtGeneralTimeOut)).BeginInit();
            this.SuspendLayout();
            // 
            // btnOKQD
            // 
            this.btnOKQD.Location = new System.Drawing.Point(175, 183);
            this.btnOKQD.Name = "btnOKQD";
            this.btnOKQD.Size = new System.Drawing.Size(124, 24);
            this.btnOKQD.TabIndex = 5;
            this.btnOKQD.Text = "OK";
            this.btnOKQD.UseVisualStyleBackColor = true;
            this.btnOKQD.Click += new System.EventHandler(this.btnOKQD_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(45, 183);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(124, 24);
            this.button1.TabIndex = 4;
            this.button1.Text = "Test connection";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // User
            // 
            this.User.Location = new System.Drawing.Point(126, 46);
            this.User.Name = "User";
            this.User.Size = new System.Drawing.Size(100, 20);
            this.User.TabIndex = 1;
            // 
            // Pass
            // 
            this.Pass.Location = new System.Drawing.Point(126, 77);
            this.Pass.Name = "Pass";
            this.Pass.PasswordChar = '*';
            this.Pass.Size = new System.Drawing.Size(100, 20);
            this.Pass.TabIndex = 2;
            this.Pass.UseSystemPasswordChar = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(22, 110);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 13);
            this.label6.TabIndex = 50;
            this.label6.Text = "Database";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(22, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(58, 13);
            this.label4.TabIndex = 49;
            this.label4.Text = "User name";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 82);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 48;
            this.label3.Text = "Password";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 47;
            this.label2.Text = "Server";
            // 
            // Server
            // 
            this.Server.FormattingEnabled = true;
            this.Server.Location = new System.Drawing.Point(126, 14);
            this.Server.Name = "Server";
            this.Server.Size = new System.Drawing.Size(164, 21);
            this.Server.TabIndex = 0;
            this.Server.DropDown += new System.EventHandler(this.Server_DropDown);
            this.Server.Enter += new System.EventHandler(this.Server_Enter);
            // 
            // Database
            // 
            this.Database.FormattingEnabled = true;
            this.Database.Location = new System.Drawing.Point(126, 108);
            this.Database.Name = "Database";
            this.Database.Size = new System.Drawing.Size(164, 21);
            this.Database.TabIndex = 3;
            this.Database.DropDown += new System.EventHandler(this.Database_DropDown);
            this.Database.Enter += new System.EventHandler(this.Database_Enter);
            // 
            // txtGeneralTimeOut
            // 
            this.txtGeneralTimeOut.Location = new System.Drawing.Point(126, 140);
            this.txtGeneralTimeOut.Maximum = new decimal(new int[] {
            60000,
            0,
            0,
            0});
            this.txtGeneralTimeOut.Name = "txtGeneralTimeOut";
            this.txtGeneralTimeOut.Size = new System.Drawing.Size(120, 20);
            this.txtGeneralTimeOut.TabIndex = 51;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 142);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(99, 13);
            this.label1.TabIndex = 52;
            this.label1.Text = "General Timeout (s)";
            // 
            // FrmConnection
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(341, 219);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtGeneralTimeOut);
            this.Controls.Add(this.Database);
            this.Controls.Add(this.Server);
            this.Controls.Add(this.User);
            this.Controls.Add(this.Pass);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnOKQD);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmConnection";
            this.Text = "Connection";
            this.Load += new System.EventHandler(this.FrmConnection_Load);
            ((System.ComponentModel.ISupportInitialize)(this.txtGeneralTimeOut)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOKQD;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox User;
        private System.Windows.Forms.TextBox Pass;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox Server;
        private System.Windows.Forms.ComboBox Database;
        private System.Windows.Forms.NumericUpDown txtGeneralTimeOut;
        private System.Windows.Forms.Label label1;
    }
}