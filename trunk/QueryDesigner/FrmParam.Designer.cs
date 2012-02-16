namespace QueryDesigner
{
    partial class FrmParam
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
            this.txtMain = new System.Windows.Forms.RichTextBox();
            this.txtName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.rdText = new System.Windows.Forms.RadioButton();
            this.rdNum = new System.Windows.Forms.RadioButton();
            this.rdDate = new System.Windows.Forms.RadioButton();
            this.btSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtMain
            // 
            this.txtMain.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtMain.EnableAutoDragDrop = true;
            this.txtMain.Location = new System.Drawing.Point(0, 0);
            this.txtMain.Name = "txtMain";
            this.txtMain.Size = new System.Drawing.Size(384, 183);
            this.txtMain.TabIndex = 8;
            this.txtMain.Text = "";
            // 
            // txtName
            // 
            this.txtName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.txtName.Location = new System.Drawing.Point(93, 196);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(283, 20);
            this.txtName.TabIndex = 10;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(51, 199);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "Name";
            // 
            // rdText
            // 
            this.rdText.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.rdText.AutoSize = true;
            this.rdText.Location = new System.Drawing.Point(93, 222);
            this.rdText.Name = "rdText";
            this.rdText.Size = new System.Drawing.Size(46, 17);
            this.rdText.TabIndex = 12;
            this.rdText.TabStop = true;
            this.rdText.Text = "Text";
            this.rdText.UseVisualStyleBackColor = true;
            // 
            // rdNum
            // 
            this.rdNum.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.rdNum.AutoSize = true;
            this.rdNum.Location = new System.Drawing.Point(292, 222);
            this.rdNum.Name = "rdNum";
            this.rdNum.Size = new System.Drawing.Size(62, 17);
            this.rdNum.TabIndex = 13;
            this.rdNum.TabStop = true;
            this.rdNum.Text = "Number";
            this.rdNum.UseVisualStyleBackColor = true;
            // 
            // rdDate
            // 
            this.rdDate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.rdDate.AutoSize = true;
            this.rdDate.Location = new System.Drawing.Point(193, 222);
            this.rdDate.Name = "rdDate";
            this.rdDate.Size = new System.Drawing.Size(48, 17);
            this.rdDate.TabIndex = 14;
            this.rdDate.TabStop = true;
            this.rdDate.Text = "Date";
            this.rdDate.UseVisualStyleBackColor = true;
            // 
            // btSave
            // 
            this.btSave.Location = new System.Drawing.Point(174, 266);
            this.btSave.Name = "btSave";
            this.btSave.Size = new System.Drawing.Size(97, 32);
            this.btSave.TabIndex = 15;
            this.btSave.Text = "OK";
            this.btSave.UseVisualStyleBackColor = true;
            this.btSave.Click += new System.EventHandler(this.btSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(288, 266);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(97, 32);
            this.btnCancel.TabIndex = 16;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // FrmParam
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 297);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btSave);
            this.Controls.Add(this.txtMain);
            this.Controls.Add(this.rdDate);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.rdNum);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.rdText);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "FrmParam";
            this.Text = "Parameters";
            this.Load += new System.EventHandler(this.FrmUserFunc_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox txtMain;
        private System.Windows.Forms.RadioButton rdText;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btSave;
        private System.Windows.Forms.RadioButton rdDate;
        private System.Windows.Forms.RadioButton rdNum;
    }
}