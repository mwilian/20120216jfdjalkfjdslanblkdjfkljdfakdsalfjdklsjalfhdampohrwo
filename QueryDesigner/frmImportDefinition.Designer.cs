namespace dCube
{
    partial class frmImportDefinition
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
            this.components = new System.ComponentModel.Container();
            Janus.Windows.GridEX.GridEXLayout dgvField_DesignTimeLayout = new Janus.Windows.GridEX.GridEXLayout();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmImportDefinition));
            this.ddlQD = new System.Windows.Forms.ComboBox();
            this.lbCON_ID = new System.Windows.Forms.Label();
            this.txtConnect = new System.Windows.Forms.TextBox();
            this.lbCode = new System.Windows.Forms.Label();
            this.lbDescription = new System.Windows.Forms.Label();
            this.txtCode = new System.Windows.Forms.TextBox();
            this.txtDescription = new System.Windows.Forms.TextBox();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.txtCommand = new System.Windows.Forms.ToolStripTextBox();
            this.btnNew = new System.Windows.Forms.ToolStripButton();
            this.btnView = new System.Windows.Forms.ToolStripButton();
            this.btnEdit = new System.Windows.Forms.ToolStripButton();
            this.btnSave = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnDelete = new System.Windows.Forms.ToolStripButton();
            this.btnCopy = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.lbErr = new System.Windows.Forms.ToolStripLabel();
            this.panelControl = new System.Windows.Forms.Panel();
            this.bt_group = new System.Windows.Forms.PictureBox();
            this.lbgroup = new System.Windows.Forms.Label();
            this.lblGroup = new System.Windows.Forms.Label();
            this.txtDAG = new System.Windows.Forms.TextBox();
            this.txtdatabase = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtConectEx = new System.Windows.Forms.TextBox();
            this.txtLookup = new System.Windows.Forms.TextBox();
            this.lbLookup = new System.Windows.Forms.Label();
            this.panelTab = new System.Windows.Forms.Panel();
            this.btnSelectTable = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabField = new System.Windows.Forms.TabPage();
            this.dgvField = new Janus.Windows.GridEX.GridEX();
            this.bsField = new System.Windows.Forms.BindingSource(this.components);
            this.toolStrip1.SuspendLayout();
            this.panelControl.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bt_group)).BeginInit();
            this.panelTab.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabField.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvField)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bsField)).BeginInit();
            this.SuspendLayout();
            // 
            // ddlQD
            // 
            this.ddlQD.DisplayMember = "KEY";
            this.ddlQD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ddlQD.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.ddlQD.FormattingEnabled = true;
            this.ddlQD.Location = new System.Drawing.Point(83, 8);
            this.ddlQD.Name = "ddlQD";
            this.ddlQD.Size = new System.Drawing.Size(75, 21);
            this.ddlQD.TabIndex = 0;
            this.ddlQD.ValueMember = "CONTENT";
            this.ddlQD.SelectedIndexChanged += new System.EventHandler(this.ddlQD_SelectedIndexChanged);
            // 
            // lbCON_ID
            // 
            this.lbCON_ID.AutoSize = true;
            this.lbCON_ID.Location = new System.Drawing.Point(16, 11);
            this.lbCON_ID.Name = "lbCON_ID";
            this.lbCON_ID.Size = new System.Drawing.Size(67, 13);
            this.lbCON_ID.TabIndex = 1;
            this.lbCON_ID.Text = "Connection";
            // 
            // txtConnect
            // 
            this.txtConnect.BackColor = System.Drawing.Color.White;
            this.txtConnect.ForeColor = System.Drawing.Color.Black;
            this.txtConnect.Location = new System.Drawing.Point(160, 29);
            this.txtConnect.Name = "txtConnect";
            this.txtConnect.ReadOnly = true;
            this.txtConnect.Size = new System.Drawing.Size(446, 22);
            this.txtConnect.TabIndex = 2;
            this.txtConnect.Visible = false;
            this.txtConnect.TextChanged += new System.EventHandler(this.txtConnect_TextChanged);
            // 
            // lbCode
            // 
            this.lbCode.AutoSize = true;
            this.lbCode.Location = new System.Drawing.Point(16, 55);
            this.lbCode.Name = "lbCode";
            this.lbCode.Size = new System.Drawing.Size(34, 13);
            this.lbCode.TabIndex = 3;
            this.lbCode.Text = "Code";
            // 
            // lbDescription
            // 
            this.lbDescription.AutoSize = true;
            this.lbDescription.Location = new System.Drawing.Point(16, 111);
            this.lbDescription.Name = "lbDescription";
            this.lbDescription.Size = new System.Drawing.Size(66, 13);
            this.lbDescription.TabIndex = 4;
            this.lbDescription.Text = "Description";
            // 
            // txtCode
            // 
            this.txtCode.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.txtCode.Location = new System.Drawing.Point(83, 52);
            this.txtCode.MaxLength = 15;
            this.txtCode.Name = "txtCode";
            this.txtCode.Size = new System.Drawing.Size(100, 22);
            this.txtCode.TabIndex = 5;
            this.txtCode.TextChanged += new System.EventHandler(this.txtCode_TextChanged);
            this.txtCode.Leave += new System.EventHandler(this.txtCode_Leave);
            // 
            // txtDescription
            // 
            this.txtDescription.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.txtDescription.Location = new System.Drawing.Point(83, 108);
            this.txtDescription.Name = "txtDescription";
            this.txtDescription.Size = new System.Drawing.Size(379, 22);
            this.txtDescription.TabIndex = 8;
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
            this.btnCopy,
            this.toolStripSeparator2,
            this.lbErr});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1042, 39);
            this.toolStrip1.TabIndex = 7;
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
            this.btnNew.Image = global::dCube.Properties.Resources.app_32x32;
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
            this.btnView.Image = global::dCube.Properties.Resources.app_search_32x32;
            this.btnView.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnView.Name = "btnView";
            this.btnView.Size = new System.Drawing.Size(36, 36);
            this.btnView.Text = "View";
            this.btnView.Click += new System.EventHandler(this.btnView_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnEdit.Image = global::dCube.Properties.Resources.app_edit_32x32;
            this.btnEdit.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(36, 36);
            this.btnEdit.Text = "Edit";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnSave
            // 
            this.btnSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnSave.Image = global::dCube.Properties.Resources.save_48x48;
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
            this.btnDelete.Image = global::dCube.Properties.Resources.app_delete_32x32;
            this.btnDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(36, 36);
            this.btnDelete.Text = "Delete";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnCopy
            // 
            this.btnCopy.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnCopy.Image = global::dCube.Properties.Resources.Copy;
            this.btnCopy.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(36, 36);
            this.btnCopy.Text = "Copy";
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 39);
            // 
            // lbErr
            // 
            this.lbErr.AutoSize = false;
            this.lbErr.Name = "lbErr";
            this.lbErr.Size = new System.Drawing.Size(100, 36);
            this.lbErr.Text = "...";
            this.lbErr.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lbErr.Click += new System.EventHandler(this.lbErr_Click);
            // 
            // panelControl
            // 
            this.panelControl.Controls.Add(this.bt_group);
            this.panelControl.Controls.Add(this.lbgroup);
            this.panelControl.Controls.Add(this.lblGroup);
            this.panelControl.Controls.Add(this.txtDAG);
            this.panelControl.Controls.Add(this.txtdatabase);
            this.panelControl.Controls.Add(this.label1);
            this.panelControl.Controls.Add(this.txtConectEx);
            this.panelControl.Controls.Add(this.txtLookup);
            this.panelControl.Controls.Add(this.lbLookup);
            this.panelControl.Controls.Add(this.lbCON_ID);
            this.panelControl.Controls.Add(this.ddlQD);
            this.panelControl.Controls.Add(this.txtDescription);
            this.panelControl.Controls.Add(this.txtConnect);
            this.panelControl.Controls.Add(this.txtCode);
            this.panelControl.Controls.Add(this.lbCode);
            this.panelControl.Controls.Add(this.lbDescription);
            this.panelControl.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelControl.Location = new System.Drawing.Point(0, 39);
            this.panelControl.Name = "panelControl";
            this.panelControl.Size = new System.Drawing.Size(1042, 139);
            this.panelControl.TabIndex = 8;
            // 
            // bt_group
            // 
            this.bt_group.BackColor = System.Drawing.Color.Transparent;
            this.bt_group.Image = global::dCube.Properties.Resources._1303882176_search_16;
            this.bt_group.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.bt_group.Location = new System.Drawing.Point(379, 83);
            this.bt_group.Name = "bt_group";
            this.bt_group.Size = new System.Drawing.Size(16, 16);
            this.bt_group.TabIndex = 52;
            this.bt_group.TabStop = false;
            this.bt_group.Click += new System.EventHandler(this.bt_group_Click);
            // 
            // lbgroup
            // 
            this.lbgroup.AutoSize = true;
            this.lbgroup.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.lbgroup.Location = new System.Drawing.Point(412, 83);
            this.lbgroup.Name = "lbgroup";
            this.lbgroup.Size = new System.Drawing.Size(22, 13);
            this.lbgroup.TabIndex = 51;
            this.lbgroup.Text = "___";
            // 
            // lblGroup
            // 
            this.lblGroup.AutoSize = true;
            this.lblGroup.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.lblGroup.Location = new System.Drawing.Point(224, 83);
            this.lblGroup.Name = "lblGroup";
            this.lblGroup.Size = new System.Drawing.Size(40, 13);
            this.lblGroup.TabIndex = 50;
            this.lblGroup.Text = "Group";
            // 
            // txtDAG
            // 
            this.txtDAG.Location = new System.Drawing.Point(285, 80);
            this.txtDAG.MaxLength = 5;
            this.txtDAG.Name = "txtDAG";
            this.txtDAG.Size = new System.Drawing.Size(88, 22);
            this.txtDAG.TabIndex = 49;
            this.txtDAG.TextChanged += new System.EventHandler(this.txtDAG_TextChanged);
            // 
            // txtdatabase
            // 
            this.txtdatabase.Location = new System.Drawing.Point(285, 52);
            this.txtdatabase.Name = "txtdatabase";
            this.txtdatabase.ReadOnly = true;
            this.txtdatabase.Size = new System.Drawing.Size(88, 22);
            this.txtdatabase.TabIndex = 46;
            this.txtdatabase.TextChanged += new System.EventHandler(this.txtdatabase_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(224, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "Database";
            // 
            // txtConectEx
            // 
            this.txtConectEx.BackColor = System.Drawing.Color.White;
            this.txtConectEx.ForeColor = System.Drawing.Color.Black;
            this.txtConectEx.Location = new System.Drawing.Point(160, 8);
            this.txtConectEx.Name = "txtConectEx";
            this.txtConectEx.ReadOnly = true;
            this.txtConectEx.Size = new System.Drawing.Size(446, 22);
            this.txtConectEx.TabIndex = 11;
            // 
            // txtLookup
            // 
            this.txtLookup.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.txtLookup.Location = new System.Drawing.Point(83, 80);
            this.txtLookup.Name = "txtLookup";
            this.txtLookup.ReadOnly = true;
            this.txtLookup.Size = new System.Drawing.Size(135, 22);
            this.txtLookup.TabIndex = 7;
            // 
            // lbLookup
            // 
            this.lbLookup.AutoSize = true;
            this.lbLookup.Location = new System.Drawing.Point(16, 83);
            this.lbLookup.Name = "lbLookup";
            this.lbLookup.Size = new System.Drawing.Size(49, 13);
            this.lbLookup.TabIndex = 7;
            this.lbLookup.Text = "Look up";
            // 
            // panelTab
            // 
            this.panelTab.Controls.Add(this.btnSelectTable);
            this.panelTab.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelTab.Location = new System.Drawing.Point(0, 471);
            this.panelTab.Name = "panelTab";
            this.panelTab.Size = new System.Drawing.Size(1042, 32);
            this.panelTab.TabIndex = 10;
            // 
            // btnSelectTable
            // 
            this.btnSelectTable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectTable.Image = global::dCube.Properties.Resources._1303882176_search_16;
            this.btnSelectTable.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSelectTable.Location = new System.Drawing.Point(942, 3);
            this.btnSelectTable.Name = "btnSelectTable";
            this.btnSelectTable.Size = new System.Drawing.Size(88, 23);
            this.btnSelectTable.TabIndex = 0;
            this.btnSelectTable.Text = "Select Table";
            this.btnSelectTable.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnSelectTable.UseVisualStyleBackColor = true;
            this.btnSelectTable.Click += new System.EventHandler(this.btnSelectTable_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.Buttons;
            this.tabControl1.Controls.Add(this.tabField);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 178);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1042, 293);
            this.tabControl1.TabIndex = 9;
            // 
            // tabField
            // 
            this.tabField.Controls.Add(this.dgvField);
            this.tabField.Location = new System.Drawing.Point(4, 25);
            this.tabField.Name = "tabField";
            this.tabField.Padding = new System.Windows.Forms.Padding(3);
            this.tabField.Size = new System.Drawing.Size(1034, 264);
            this.tabField.TabIndex = 1;
            this.tabField.Text = "Field";
            this.tabField.UseVisualStyleBackColor = true;
            // 
            // dgvField
            // 
            this.dgvField.AllowAddNew = Janus.Windows.GridEX.InheritableBoolean.True;
            this.dgvField.AllowDelete = Janus.Windows.GridEX.InheritableBoolean.True;
            this.dgvField.AllowDrop = true;
            this.dgvField.ColumnAutoSizeMode = Janus.Windows.GridEX.ColumnAutoSizeMode.DisplayedCellsAndHeader;
            this.dgvField.DataSource = this.bsField;
            dgvField_DesignTimeLayout.LayoutString = resources.GetString("dgvField_DesignTimeLayout.LayoutString");
            this.dgvField.DesignTimeLayout = dgvField_DesignTimeLayout;
            this.dgvField.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvField.Font = new System.Drawing.Font("Segoe UI", 8.25F);
            this.dgvField.GroupByBoxVisible = false;
            this.dgvField.Location = new System.Drawing.Point(3, 3);
            this.dgvField.Name = "dgvField";
            this.dgvField.RowHeaders = Janus.Windows.GridEX.InheritableBoolean.True;
            this.dgvField.Size = new System.Drawing.Size(1028, 258);
            this.dgvField.TabIndex = 1;
            this.dgvField.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            this.dgvField.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dgvField_MouseUp);
            this.dgvField.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dgvField_MouseDown);
            this.dgvField.LinkClicked += new Janus.Windows.GridEX.ColumnActionEventHandler(this.dgvField_LinkClicked);
            this.dgvField.DragEnter += new System.Windows.Forms.DragEventHandler(this.dgvField_DragEnter);
            this.dgvField.DragDrop += new System.Windows.Forms.DragEventHandler(this.dgvField_DragDrop);
            // 
            // frmImportDefinition
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1042, 503);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.panelTab);
            this.Controls.Add(this.panelControl);
            this.Controls.Add(this.toolStrip1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "frmImportDefinition";
            this.Text = "Import Definition";
            this.Load += new System.EventHandler(this.frmQDADD_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.panelControl.ResumeLayout(false);
            this.panelControl.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bt_group)).EndInit();
            this.panelTab.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabField.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvField)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bsField)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox ddlQD;
        private System.Windows.Forms.Label lbCON_ID;
        private System.Windows.Forms.TextBox txtConnect;
        private System.Windows.Forms.Label lbCode;
        private System.Windows.Forms.Label lbDescription;
        private System.Windows.Forms.TextBox txtCode;
        private System.Windows.Forms.TextBox txtDescription;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripTextBox txtCommand;
        private System.Windows.Forms.ToolStripButton btnNew;
        private System.Windows.Forms.ToolStripButton btnView;
        private System.Windows.Forms.ToolStripButton btnEdit;
        private System.Windows.Forms.ToolStripButton btnSave;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton btnDelete;
        private System.Windows.Forms.ToolStripButton btnCopy;
        private System.Windows.Forms.Panel panelControl;
        private System.Windows.Forms.Panel panelTab;
        private System.Windows.Forms.Button btnSelectTable;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabField;
        private System.Windows.Forms.BindingSource bsField;
        private System.Windows.Forms.TextBox txtLookup;
        private System.Windows.Forms.Label lbLookup;
        private Janus.Windows.GridEX.GridEX dgvField;
        private System.Windows.Forms.TextBox txtConectEx;
        private System.Windows.Forms.ToolStripLabel lbErr;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtdatabase;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.PictureBox bt_group;
        private System.Windows.Forms.Label lbgroup;
        private System.Windows.Forms.Label lblGroup;
        private System.Windows.Forms.TextBox txtDAG;
    }
}