namespace QueryDesigner
{
    partial class FrmLicense
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtNumUser = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtSerial = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtCompany = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ckbTask = new System.Windows.Forms.CheckBox();
            this.ckbWeb = new System.Windows.Forms.CheckBox();
            this.ckbAddin = new System.Windows.Forms.CheckBox();
            this.ckbQDADD = new System.Windows.Forms.CheckBox();
            this.ckbQD = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.lbErr = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.dtExpiryDate = new System.Windows.Forms.DateTimePicker();
            this.txtKey = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtNumUser);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtSerial);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.txtCompany);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(504, 111);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Information";
            // 
            // txtNumUser
            // 
            this.txtNumUser.Location = new System.Drawing.Point(94, 72);
            this.txtNumUser.Name = "txtNumUser";
            this.txtNumUser.Size = new System.Drawing.Size(92, 20);
            this.txtNumUser.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 75);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(81, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Number of User";
            // 
            // txtSerial
            // 
            this.txtSerial.Location = new System.Drawing.Point(94, 46);
            this.txtSerial.Name = "txtSerial";
            this.txtSerial.Size = new System.Drawing.Size(128, 20);
            this.txtSerial.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(33, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Serial";
            // 
            // txtCompany
            // 
            this.txtCompany.Location = new System.Drawing.Point(94, 20);
            this.txtCompany.Name = "txtCompany";
            this.txtCompany.Size = new System.Drawing.Size(404, 20);
            this.txtCompany.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Company Name";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.ckbTask);
            this.groupBox2.Controls.Add(this.ckbWeb);
            this.groupBox2.Controls.Add(this.ckbAddin);
            this.groupBox2.Controls.Add(this.ckbQDADD);
            this.groupBox2.Controls.Add(this.ckbQD);
            this.groupBox2.Location = new System.Drawing.Point(12, 129);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(504, 74);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Serialized Modules";
            // 
            // ckbTask
            // 
            this.ckbTask.AutoSize = true;
            this.ckbTask.Location = new System.Drawing.Point(240, 19);
            this.ckbTask.Name = "ckbTask";
            this.ckbTask.Size = new System.Drawing.Size(47, 17);
            this.ckbTask.TabIndex = 4;
            this.ckbTask.Text = "Alert";
            this.ckbTask.UseVisualStyleBackColor = true;
            // 
            // ckbWeb
            // 
            this.ckbWeb.AutoSize = true;
            this.ckbWeb.Location = new System.Drawing.Point(24, 42);
            this.ckbWeb.Name = "ckbWeb";
            this.ckbWeb.Size = new System.Drawing.Size(92, 17);
            this.ckbWeb.TabIndex = 1;
            this.ckbWeb.Text = "Web preiewer";
            this.ckbWeb.UseVisualStyleBackColor = true;
            // 
            // ckbAddin
            // 
            this.ckbAddin.AutoSize = true;
            this.ckbAddin.Location = new System.Drawing.Point(129, 19);
            this.ckbAddin.Name = "ckbAddin";
            this.ckbAddin.Size = new System.Drawing.Size(53, 17);
            this.ckbAddin.TabIndex = 2;
            this.ckbAddin.Text = "Addin";
            this.ckbAddin.UseVisualStyleBackColor = true;
            // 
            // ckbQDADD
            // 
            this.ckbQDADD.AutoSize = true;
            this.ckbQDADD.Location = new System.Drawing.Point(129, 42);
            this.ckbQDADD.Name = "ckbQDADD";
            this.ckbQDADD.Size = new System.Drawing.Size(113, 17);
            this.ckbQDADD.TabIndex = 3;
            this.ckbQDADD.Text = "Dictionary Building";
            this.ckbQDADD.UseVisualStyleBackColor = true;
            // 
            // ckbQD
            // 
            this.ckbQD.AutoSize = true;
            this.ckbQD.Location = new System.Drawing.Point(24, 19);
            this.ckbQD.Name = "ckbQD";
            this.ckbQD.Size = new System.Drawing.Size(99, 17);
            this.ckbQD.TabIndex = 0;
            this.ckbQD.Text = "Query Designer";
            this.ckbQD.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lbErr);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.dtExpiryDate);
            this.groupBox3.Controls.Add(this.txtKey);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Location = new System.Drawing.Point(12, 209);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(504, 109);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Key";
            // 
            // lbErr
            // 
            this.lbErr.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.lbErr.AutoSize = true;
            this.lbErr.ForeColor = System.Drawing.Color.Red;
            this.lbErr.Location = new System.Drawing.Point(239, 73);
            this.lbErr.Name = "lbErr";
            this.lbErr.Size = new System.Drawing.Size(35, 13);
            this.lbErr.TabIndex = 44;
            this.lbErr.Text = "label6";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(20, 25);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(61, 13);
            this.label5.TabIndex = 43;
            this.label5.Text = "Expiry Date";
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // dtExpiryDate
            // 
            this.dtExpiryDate.CustomFormat = "dd/MM/yyyy";
            this.dtExpiryDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtExpiryDate.Location = new System.Drawing.Point(93, 19);
            this.dtExpiryDate.Name = "dtExpiryDate";
            this.dtExpiryDate.Size = new System.Drawing.Size(108, 20);
            this.dtExpiryDate.TabIndex = 0;
            // 
            // txtKey
            // 
            this.txtKey.Location = new System.Drawing.Point(93, 50);
            this.txtKey.Name = "txtKey";
            this.txtKey.Size = new System.Drawing.Size(404, 20);
            this.txtKey.TabIndex = 1;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(21, 60);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(25, 13);
            this.label4.TabIndex = 2;
            this.label4.Text = "Key";
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(136, 327);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(110, 24);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(252, 327);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(110, 24);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // FrmLicense
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(535, 374);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.MaximumSize = new System.Drawing.Size(543, 401);
            this.MinimumSize = new System.Drawing.Size(543, 401);
            this.Name = "FrmLicense";
            this.Text = "Registry License";
            this.Load += new System.EventHandler(this.FrmLicense_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox txtCompany;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox txtNumUser;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtSerial;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox ckbWeb;
        private System.Windows.Forms.CheckBox ckbAddin;
        private System.Windows.Forms.CheckBox ckbQDADD;
        private System.Windows.Forms.CheckBox ckbQD;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DateTimePicker dtExpiryDate;
        private System.Windows.Forms.TextBox txtKey;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lbErr;
        private System.Windows.Forms.CheckBox ckbTask;
    }
}