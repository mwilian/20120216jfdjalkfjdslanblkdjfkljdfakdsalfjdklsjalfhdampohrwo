namespace QueryDesigner
{
    partial class frmPOG
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
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.txtCommand = new System.Windows.Forms.ToolStripTextBox();
            this.btnNew = new System.Windows.Forms.ToolStripButton();
            this.btnView = new System.Windows.Forms.ToolStripButton();
            this.btnEdit = new System.Windows.Forms.ToolStripButton();
            this.btnSave = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnDelete = new System.Windows.Forms.ToolStripButton();
            this.btnCopy = new System.Windows.Forms.ToolStripButton();
            this.label1 = new System.Windows.Forms.Label();
            this.txtCode = new System.Windows.Forms.TextBox();
            this.txtName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtLen = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtValid = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtQD = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.lbErr = new System.Windows.Forms.ToolStripStatusLabel();
            this.btnDB = new System.Windows.Forms.PictureBox();
            this.pContain = new System.Windows.Forms.Panel();
            this.lbDB = new System.Windows.Forms.Label();
            this.toolStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnDB)).BeginInit();
            this.pContain.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.txtCommand,
            this.btnNew,
            this.btnView,
            this.btnEdit,
            this.btnSave,
            this.toolStripSeparator1,
            this.btnDelete,
            this.btnCopy});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(508, 39);
            this.toolStrip1.TabIndex = 8;
            this.toolStrip1.Text = "Export Excel";
            // 
            // txtCommand
            // 
            this.txtCommand.Name = "txtCommand";
            this.txtCommand.Size = new System.Drawing.Size(100, 39);
            // 
            // btnNew
            // 
            this.btnNew.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnNew.Image = global::QueryDesigner.Properties.Resources.app_32x32;
            this.btnNew.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(36, 36);
            this.btnNew.Text = "New";
            this.btnNew.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal;
            this.btnNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnView
            // 
            this.btnView.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnView.Image = global::QueryDesigner.Properties.Resources.app_search_32x32;
            this.btnView.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnView.Name = "btnView";
            this.btnView.Size = new System.Drawing.Size(36, 36);
            this.btnView.Text = "View";
            this.btnView.Click += new System.EventHandler(this.btnView_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnEdit.Image = global::QueryDesigner.Properties.Resources.app_edit_32x32;
            this.btnEdit.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(36, 36);
            this.btnEdit.Text = "Edit";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnSave
            // 
            this.btnSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnSave.Image = global::QueryDesigner.Properties.Resources.save_48x48;
            this.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(36, 36);
            this.btnSave.Text = "Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 39);
            // 
            // btnDelete
            // 
            this.btnDelete.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnDelete.Image = global::QueryDesigner.Properties.Resources.app_delete_32x32;
            this.btnDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(36, 36);
            this.btnDelete.Text = "Delete";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnCopy
            // 
            this.btnCopy.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnCopy.Image = global::QueryDesigner.Properties.Resources.Copy;
            this.btnCopy.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(36, 36);
            this.btnCopy.Text = "Copy";
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Group Code";
            // 
            // txtCode
            // 
            this.txtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtCode.Location = new System.Drawing.Point(102, 12);
            this.txtCode.MaxLength = 5;
            this.txtCode.Name = "txtCode";
            this.txtCode.Size = new System.Drawing.Size(100, 20);
            this.txtCode.TabIndex = 10;
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(102, 38);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(271, 20);
            this.txtName.TabIndex = 12;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Group Name";
            // 
            // txtLen
            // 
            this.txtLen.Location = new System.Drawing.Point(154, 64);
            this.txtLen.Name = "txtLen";
            this.txtLen.Size = new System.Drawing.Size(48, 20);
            this.txtLen.TabIndex = 14;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 67);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(131, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Password Minimun Length";
            // 
            // txtValid
            // 
            this.txtValid.Location = new System.Drawing.Point(102, 90);
            this.txtValid.Name = "txtValid";
            this.txtValid.Size = new System.Drawing.Size(100, 20);
            this.txtValid.TabIndex = 16;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 93);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(79, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Password Valid";
            // 
            // txtQD
            // 
            this.txtQD.Location = new System.Drawing.Point(116, 116);
            this.txtQD.Name = "txtQD";
            this.txtQD.Size = new System.Drawing.Size(86, 20);
            this.txtQD.TabIndex = 18;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(17, 119);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(39, 13);
            this.label5.TabIndex = 17;
            this.label5.Text = "Report";
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lbErr});
            this.statusStrip1.Location = new System.Drawing.Point(0, 278);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(508, 22);
            this.statusStrip1.TabIndex = 19;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // lbErr
            // 
            this.lbErr.AutoSize = false;
            this.lbErr.Name = "lbErr";
            this.lbErr.Size = new System.Drawing.Size(200, 17);
            this.lbErr.Text = "...";
            this.lbErr.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lbErr.Click += new System.EventHandler(this.lbErr_Click);
            // 
            // btnDB
            // 
            this.btnDB.BackColor = System.Drawing.Color.Transparent;
            this.btnDB.Image = global::QueryDesigner.Properties.Resources._1303882176_search_16;
            this.btnDB.Location = new System.Drawing.Point(208, 119);
            this.btnDB.Name = "btnDB";
            this.btnDB.Size = new System.Drawing.Size(16, 16);
            this.btnDB.TabIndex = 49;
            this.btnDB.TabStop = false;
            this.btnDB.Click += new System.EventHandler(this.btnDB_Click);
            // 
            // pContain
            // 
            this.pContain.Controls.Add(this.lbDB);
            this.pContain.Controls.Add(this.txtCode);
            this.pContain.Controls.Add(this.btnDB);
            this.pContain.Controls.Add(this.label1);
            this.pContain.Controls.Add(this.label2);
            this.pContain.Controls.Add(this.txtName);
            this.pContain.Controls.Add(this.txtQD);
            this.pContain.Controls.Add(this.label3);
            this.pContain.Controls.Add(this.label5);
            this.pContain.Controls.Add(this.txtLen);
            this.pContain.Controls.Add(this.txtValid);
            this.pContain.Controls.Add(this.label4);
            this.pContain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pContain.Location = new System.Drawing.Point(0, 39);
            this.pContain.Name = "pContain";
            this.pContain.Size = new System.Drawing.Size(508, 239);
            this.pContain.TabIndex = 50;
            // 
            // lbDB
            // 
            this.lbDB.AutoSize = true;
            this.lbDB.Location = new System.Drawing.Point(230, 119);
            this.lbDB.Name = "lbDB";
            this.lbDB.Size = new System.Drawing.Size(13, 13);
            this.lbDB.TabIndex = 51;
            this.lbDB.Text = "_";
            // 
            // frmPOG
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(508, 300);
            this.Controls.Add(this.pContain);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.toolStrip1);
            this.Name = "frmPOG";
            this.Text = "User Definition";
            this.Load += new System.EventHandler(this.frmPOD_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnDB)).EndInit();
            this.pContain.ResumeLayout(false);
            this.pContain.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripTextBox txtCommand;
        private System.Windows.Forms.ToolStripButton btnNew;
        private System.Windows.Forms.ToolStripButton btnView;
        private System.Windows.Forms.ToolStripButton btnEdit;
        private System.Windows.Forms.ToolStripButton btnSave;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton btnDelete;
        private System.Windows.Forms.ToolStripButton btnCopy;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtCode;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtLen;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtValid;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtQD;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel lbErr;
        private System.Windows.Forms.PictureBox btnDB;
        private System.Windows.Forms.Panel pContain;
        private System.Windows.Forms.Label lbDB;
    }
}