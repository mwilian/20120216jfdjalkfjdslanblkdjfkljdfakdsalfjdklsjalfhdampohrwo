namespace dCube
{
    partial class FrmUserFunc
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
            Janus.Windows.GridEX.GridEXLayout dgvSelectNodes_DesignTimeLayout = new Janus.Windows.GridEX.GridEXLayout();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmUserFunc));
            this.txtMain = new System.Windows.Forms.RichTextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btSave = new System.Windows.Forms.Button();
            this.rdDate = new System.Windows.Forms.RadioButton();
            this.txtName = new System.Windows.Forms.TextBox();
            this.rdNum = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.rdText = new System.Windows.Forms.RadioButton();
            this.btnNew = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.dgvSelectNodes = new Janus.Windows.GridEX.GridEX();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSelectNodes)).BeginInit();
            this.SuspendLayout();
            // 
            // txtMain
            // 
            this.txtMain.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtMain.EnableAutoDragDrop = true;
            this.txtMain.Location = new System.Drawing.Point(0, 0);
            this.txtMain.Name = "txtMain";
            this.txtMain.Size = new System.Drawing.Size(378, 282);
            this.txtMain.TabIndex = 8;
            this.txtMain.Text = "";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(276, 358);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(97, 32);
            this.btnCancel.TabIndex = 23;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btSave_Click);
            // 
            // btSave
            // 
            this.btSave.Location = new System.Drawing.Point(173, 358);
            this.btSave.Name = "btSave";
            this.btSave.Size = new System.Drawing.Size(97, 32);
            this.btSave.TabIndex = 22;
            this.btSave.Text = "OK";
            this.btSave.UseVisualStyleBackColor = true;
            this.btSave.Click += new System.EventHandler(this.btSave_Click);
            // 
            // rdDate
            // 
            this.rdDate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.rdDate.AutoSize = true;
            this.rdDate.Location = new System.Drawing.Point(182, 314);
            this.rdDate.Name = "rdDate";
            this.rdDate.Size = new System.Drawing.Size(48, 17);
            this.rdDate.TabIndex = 21;
            this.rdDate.TabStop = true;
            this.rdDate.Text = "Date";
            this.rdDate.UseVisualStyleBackColor = true;
            // 
            // txtName
            // 
            this.txtName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.txtName.Location = new System.Drawing.Point(83, 288);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(283, 20);
            this.txtName.TabIndex = 17;
            // 
            // rdNum
            // 
            this.rdNum.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.rdNum.AutoSize = true;
            this.rdNum.Location = new System.Drawing.Point(278, 314);
            this.rdNum.Name = "rdNum";
            this.rdNum.Size = new System.Drawing.Size(62, 17);
            this.rdNum.TabIndex = 20;
            this.rdNum.TabStop = true;
            this.rdNum.Text = "Number";
            this.rdNum.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(41, 291);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 18;
            this.label1.Text = "_Name";
            // 
            // rdText
            // 
            this.rdText.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.rdText.AutoSize = true;
            this.rdText.Location = new System.Drawing.Point(84, 314);
            this.rdText.Name = "rdText";
            this.rdText.Size = new System.Drawing.Size(46, 17);
            this.rdText.TabIndex = 19;
            this.rdText.TabStop = true;
            this.rdText.Text = "Text";
            this.rdText.UseVisualStyleBackColor = true;
            // 
            // btnNew
            // 
            this.btnNew.Location = new System.Drawing.Point(70, 358);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(97, 32);
            this.btnNew.TabIndex = 24;
            this.btnNew.Text = "Refresh";
            this.btnNew.UseVisualStyleBackColor = true;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnNew);
            this.panel1.Controls.Add(this.txtMain);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.txtName);
            this.panel1.Controls.Add(this.btSave);
            this.panel1.Controls.Add(this.rdText);
            this.panel1.Controls.Add(this.rdDate);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.rdNum);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel1.Location = new System.Drawing.Point(415, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(378, 396);
            this.panel1.TabIndex = 8;
            // 
            // dgvSelectNodes
            // 
            this.dgvSelectNodes.AllowDrop = true;
            this.dgvSelectNodes.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.False;
            dgvSelectNodes_DesignTimeLayout.LayoutString = resources.GetString("dgvSelectNodes_DesignTimeLayout.LayoutString");
            this.dgvSelectNodes.DesignTimeLayout = dgvSelectNodes_DesignTimeLayout;
            this.dgvSelectNodes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvSelectNodes.GroupByBoxVisible = false;
            this.dgvSelectNodes.Location = new System.Drawing.Point(0, 0);
            this.dgvSelectNodes.Name = "dgvSelectNodes";
            this.dgvSelectNodes.Size = new System.Drawing.Size(415, 396);
            this.dgvSelectNodes.TabIndex = 9;
            this.dgvSelectNodes.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            this.dgvSelectNodes.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dgvSelectNodes_MouseDown);
            // 
            // FrmUserFunc
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(793, 396);
            this.Controls.Add(this.dgvSelectNodes);
            this.Controls.Add(this.panel1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "FrmUserFunc";
            this.Text = "FrmUserFunc";
            this.Load += new System.EventHandler(this.FrmUserFunc_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSelectNodes)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox txtMain;
        private System.Windows.Forms.Button btnNew;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btSave;
        private System.Windows.Forms.RadioButton rdDate;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.RadioButton rdNum;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton rdText;
        private System.Windows.Forms.Panel panel1;
        private Janus.Windows.GridEX.GridEX dgvSelectNodes;
    }
}