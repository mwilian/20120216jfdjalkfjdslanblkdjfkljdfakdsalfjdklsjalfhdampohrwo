namespace QueryDesigner
{
    partial class frmQDADD
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
            Janus.Windows.GridEX.GridEXLayout dgvFrom_DesignTimeLayout = new Janus.Windows.GridEX.GridEXLayout();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmQDADD));
            Janus.Windows.GridEX.GridEXLayout dgvField_DesignTimeLayout = new Janus.Windows.GridEX.GridEXLayout();
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
            this.btnTransferIn = new System.Windows.Forms.ToolStripButton();
            this.btnTransferOut = new System.Windows.Forms.ToolStripButton();
            this.lbErr = new System.Windows.Forms.ToolStripLabel();
            this.panelControl = new System.Windows.Forms.Panel();
            this.btnXML = new System.Windows.Forms.Button();
            this.txtConectEx = new System.Windows.Forms.TextBox();
            this.txtModule = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.ckbUse = new System.Windows.Forms.CheckBox();
            this.txtLookup = new System.Windows.Forms.TextBox();
            this.lbLookup = new System.Windows.Forms.Label();
            this.panelTab = new System.Windows.Forms.Panel();
            this.btnRelation = new System.Windows.Forms.Button();
            this.btnSelectView = new System.Windows.Forms.Button();
            this.btnSelectTable = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabTable = new System.Windows.Forms.TabPage();
            this.dgvFrom = new Janus.Windows.GridEX.GridEX();
            this.bsFROMCODE = new System.Windows.Forms.BindingSource(this.components);
            this.tabField = new System.Windows.Forms.TabPage();
            this.dgvField = new Janus.Windows.GridEX.GridEX();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.dgvAddRow = new System.Windows.Forms.ToolStripMenuItem();
            this.bsField = new System.Windows.Forms.BindingSource(this.components);
            this.Group = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnQD = new System.Windows.Forms.PictureBox();
            this.toolStrip1.SuspendLayout();
            this.panelControl.SuspendLayout();
            this.panelTab.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabTable.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFrom)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bsFROMCODE)).BeginInit();
            this.tabField.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvField)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bsField)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnQD)).BeginInit();
            this.SuspendLayout();
            // 
            // ddlQD
            // 
            this.ddlQD.DisplayMember = "KEY";
            this.ddlQD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ddlQD.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.ddlQD.FormattingEnabled = true;
            this.ddlQD.Location = new System.Drawing.Point(91, 87);
            this.ddlQD.Name = "ddlQD";
            this.ddlQD.Size = new System.Drawing.Size(75, 21);
            this.ddlQD.TabIndex = 0;
            this.ddlQD.ValueMember = "CONTENT";
            this.ddlQD.SelectedIndexChanged += new System.EventHandler(this.ddlQD_SelectedIndexChanged);
            // 
            // lbCON_ID
            // 
            this.lbCON_ID.AutoSize = true;
            this.lbCON_ID.Location = new System.Drawing.Point(24, 90);
            this.lbCON_ID.Name = "lbCON_ID";
            this.lbCON_ID.Size = new System.Drawing.Size(67, 13);
            this.lbCON_ID.TabIndex = 1;
            this.lbCON_ID.Text = "Connection";
            // 
            // txtConnect
            // 
            this.txtConnect.BackColor = System.Drawing.Color.White;
            this.txtConnect.ForeColor = System.Drawing.Color.Black;
            this.txtConnect.Location = new System.Drawing.Point(367, 36);
            this.txtConnect.Name = "txtConnect";
            this.txtConnect.ReadOnly = true;
            this.txtConnect.Size = new System.Drawing.Size(302, 22);
            this.txtConnect.TabIndex = 2;
            this.txtConnect.Visible = false;
            this.txtConnect.TextChanged += new System.EventHandler(this.txtConnect_TextChanged);
            // 
            // lbCode
            // 
            this.lbCode.AutoSize = true;
            this.lbCode.Location = new System.Drawing.Point(24, 6);
            this.lbCode.Name = "lbCode";
            this.lbCode.Size = new System.Drawing.Size(34, 13);
            this.lbCode.TabIndex = 3;
            this.lbCode.Text = "Code";
            // 
            // lbDescription
            // 
            this.lbDescription.AutoSize = true;
            this.lbDescription.Location = new System.Drawing.Point(24, 62);
            this.lbDescription.Name = "lbDescription";
            this.lbDescription.Size = new System.Drawing.Size(66, 13);
            this.lbDescription.TabIndex = 4;
            this.lbDescription.Text = "Description";
            // 
            // txtCode
            // 
            this.txtCode.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.txtCode.Location = new System.Drawing.Point(91, 3);
            this.txtCode.MaxLength = 15;
            this.txtCode.Name = "txtCode";
            this.txtCode.Size = new System.Drawing.Size(100, 22);
            this.txtCode.TabIndex = 5;
            this.txtCode.TextChanged += new System.EventHandler(this.txtCode_TextChanged);
            // 
            // txtDescription
            // 
            this.txtDescription.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.txtDescription.Location = new System.Drawing.Point(91, 59);
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
            this.btnTransferIn,
            this.btnTransferOut,
            this.lbErr});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(912, 39);
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
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 39);
            // 
            // btnTransferIn
            // 
            this.btnTransferIn.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnTransferIn.Image = global::QueryDesigner.Properties.Resources.download;
            this.btnTransferIn.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnTransferIn.Name = "btnTransferIn";
            this.btnTransferIn.Size = new System.Drawing.Size(36, 36);
            this.btnTransferIn.Text = "Transfer In";
            this.btnTransferIn.Click += new System.EventHandler(this.btnTransferIn_Click);
            // 
            // btnTransferOut
            // 
            this.btnTransferOut.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnTransferOut.Image = global::QueryDesigner.Properties.Resources.app_upload_32x32;
            this.btnTransferOut.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnTransferOut.Name = "btnTransferOut";
            this.btnTransferOut.Size = new System.Drawing.Size(36, 36);
            this.btnTransferOut.Text = "Transfer Out";
            this.btnTransferOut.Click += new System.EventHandler(this.btnTransferOut_Click);
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
            this.panelControl.Controls.Add(this.btnQD);
            this.panelControl.Controls.Add(this.Group);
            this.panelControl.Controls.Add(this.label2);
            this.panelControl.Controls.Add(this.btnXML);
            this.panelControl.Controls.Add(this.txtConectEx);
            this.panelControl.Controls.Add(this.txtModule);
            this.panelControl.Controls.Add(this.label1);
            this.panelControl.Controls.Add(this.ckbUse);
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
            this.panelControl.Size = new System.Drawing.Size(912, 144);
            this.panelControl.TabIndex = 8;
            // 
            // btnXML
            // 
            this.btnXML.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnXML.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnXML.Location = new System.Drawing.Point(812, 110);
            this.btnXML.Name = "btnXML";
            this.btnXML.Size = new System.Drawing.Size(88, 23);
            this.btnXML.TabIndex = 12;
            this.btnXML.Text = "Import XML";
            this.btnXML.UseVisualStyleBackColor = true;
            this.btnXML.Click += new System.EventHandler(this.btnXML_Click);
            // 
            // txtConectEx
            // 
            this.txtConectEx.BackColor = System.Drawing.Color.White;
            this.txtConectEx.ForeColor = System.Drawing.Color.Black;
            this.txtConectEx.Location = new System.Drawing.Point(168, 87);
            this.txtConectEx.Name = "txtConectEx";
            this.txtConectEx.ReadOnly = true;
            this.txtConectEx.Size = new System.Drawing.Size(302, 22);
            this.txtConectEx.TabIndex = 11;
            // 
            // txtModule
            // 
            this.txtModule.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.txtModule.Location = new System.Drawing.Point(295, 3);
            this.txtModule.Name = "txtModule";
            this.txtModule.Size = new System.Drawing.Size(100, 22);
            this.txtModule.TabIndex = 10;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(228, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Module";
            // 
            // ckbUse
            // 
            this.ckbUse.AutoSize = true;
            this.ckbUse.Location = new System.Drawing.Point(295, 36);
            this.ckbUse.Name = "ckbUse";
            this.ckbUse.Size = new System.Drawing.Size(45, 17);
            this.ckbUse.TabIndex = 6;
            this.ckbUse.Text = "Use";
            this.ckbUse.UseVisualStyleBackColor = true;
            // 
            // txtLookup
            // 
            this.txtLookup.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.txtLookup.Location = new System.Drawing.Point(91, 31);
            this.txtLookup.Name = "txtLookup";
            this.txtLookup.Size = new System.Drawing.Size(135, 22);
            this.txtLookup.TabIndex = 7;
            // 
            // lbLookup
            // 
            this.lbLookup.AutoSize = true;
            this.lbLookup.Location = new System.Drawing.Point(24, 34);
            this.lbLookup.Name = "lbLookup";
            this.lbLookup.Size = new System.Drawing.Size(49, 13);
            this.lbLookup.TabIndex = 7;
            this.lbLookup.Text = "Look up";
            // 
            // panelTab
            // 
            this.panelTab.Controls.Add(this.btnRelation);
            this.panelTab.Controls.Add(this.btnSelectView);
            this.panelTab.Controls.Add(this.btnSelectTable);
            this.panelTab.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelTab.Location = new System.Drawing.Point(0, 567);
            this.panelTab.Name = "panelTab";
            this.panelTab.Size = new System.Drawing.Size(912, 32);
            this.panelTab.TabIndex = 10;
            // 
            // btnRelation
            // 
            this.btnRelation.Location = new System.Drawing.Point(12, 6);
            this.btnRelation.Name = "btnRelation";
            this.btnRelation.Size = new System.Drawing.Size(75, 23);
            this.btnRelation.TabIndex = 2;
            this.btnRelation.Text = "Relation";
            this.btnRelation.UseVisualStyleBackColor = true;
            this.btnRelation.Click += new System.EventHandler(this.btnRelation_Click);
            // 
            // btnSelectView
            // 
            this.btnSelectView.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectView.Image = global::QueryDesigner.Properties.Resources._1303882176_search_16;
            this.btnSelectView.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSelectView.Location = new System.Drawing.Point(814, 3);
            this.btnSelectView.Name = "btnSelectView";
            this.btnSelectView.Size = new System.Drawing.Size(86, 23);
            this.btnSelectView.TabIndex = 1;
            this.btnSelectView.Text = "Select View";
            this.btnSelectView.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnSelectView.UseVisualStyleBackColor = true;
            this.btnSelectView.Click += new System.EventHandler(this.btnSelectView_Click);
            // 
            // btnSelectTable
            // 
            this.btnSelectTable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectTable.Image = global::QueryDesigner.Properties.Resources._1303882176_search_16;
            this.btnSelectTable.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSelectTable.Location = new System.Drawing.Point(720, 3);
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
            this.tabControl1.Controls.Add(this.tabTable);
            this.tabControl1.Controls.Add(this.tabField);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 183);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(912, 384);
            this.tabControl1.TabIndex = 9;
            // 
            // tabTable
            // 
            this.tabTable.Controls.Add(this.dgvFrom);
            this.tabTable.Location = new System.Drawing.Point(4, 25);
            this.tabTable.Name = "tabTable";
            this.tabTable.Padding = new System.Windows.Forms.Padding(3);
            this.tabTable.Size = new System.Drawing.Size(904, 355);
            this.tabTable.TabIndex = 0;
            this.tabTable.Text = "Table & Relation";
            this.tabTable.UseVisualStyleBackColor = true;
            // 
            // dgvFrom
            // 
            this.dgvFrom.AllowAddNew = Janus.Windows.GridEX.InheritableBoolean.True;
            this.dgvFrom.AllowDelete = Janus.Windows.GridEX.InheritableBoolean.True;
            this.dgvFrom.AllowDrop = true;
            this.dgvFrom.ColumnAutoSizeMode = Janus.Windows.GridEX.ColumnAutoSizeMode.DisplayedCellsAndHeader;
            this.dgvFrom.DataSource = this.bsFROMCODE;
            dgvFrom_DesignTimeLayout.LayoutString = resources.GetString("dgvFrom_DesignTimeLayout.LayoutString");
            this.dgvFrom.DesignTimeLayout = dgvFrom_DesignTimeLayout;
            this.dgvFrom.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvFrom.GroupByBoxVisible = false;
            this.dgvFrom.Location = new System.Drawing.Point(3, 3);
            this.dgvFrom.Name = "dgvFrom";
            this.dgvFrom.RowHeaders = Janus.Windows.GridEX.InheritableBoolean.True;
            this.dgvFrom.Size = new System.Drawing.Size(898, 349);
            this.dgvFrom.TabIndex = 1;
            this.dgvFrom.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            this.dgvFrom.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dgvFrom_MouseDown);
            this.dgvFrom.DragEnter += new System.Windows.Forms.DragEventHandler(this.dgvFrom_DragEnter);
            this.dgvFrom.DragDrop += new System.Windows.Forms.DragEventHandler(this.dgvFrom_DragDrop);
            // 
            // tabField
            // 
            this.tabField.Controls.Add(this.dgvField);
            this.tabField.Location = new System.Drawing.Point(4, 25);
            this.tabField.Name = "tabField";
            this.tabField.Padding = new System.Windows.Forms.Padding(3);
            this.tabField.Size = new System.Drawing.Size(904, 360);
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
            this.dgvField.ContextMenuStrip = this.contextMenuStrip1;
            this.dgvField.DataSource = this.bsField;
            dgvField_DesignTimeLayout.LayoutString = resources.GetString("dgvField_DesignTimeLayout.LayoutString");
            this.dgvField.DesignTimeLayout = dgvField_DesignTimeLayout;
            this.dgvField.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvField.GroupByBoxVisible = false;
            this.dgvField.Location = new System.Drawing.Point(3, 3);
            this.dgvField.Name = "dgvField";
            this.dgvField.RowHeaders = Janus.Windows.GridEX.InheritableBoolean.True;
            this.dgvField.Size = new System.Drawing.Size(898, 354);
            this.dgvField.TabIndex = 1;
            this.dgvField.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            this.dgvField.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dgvField_MouseDown);
            this.dgvField.DragEnter += new System.Windows.Forms.DragEventHandler(this.dgvField_DragEnter);
            this.dgvField.DragDrop += new System.Windows.Forms.DragEventHandler(this.dgvField_DragDrop);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.dgvAddRow});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(118, 26);
            // 
            // dgvAddRow
            // 
            this.dgvAddRow.Name = "dgvAddRow";
            this.dgvAddRow.Size = new System.Drawing.Size(117, 22);
            this.dgvAddRow.Text = "Add Row";
            this.dgvAddRow.Click += new System.EventHandler(this.dgvAddRow_Click);
            // 
            // Group
            // 
            this.Group.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.Group.Location = new System.Drawing.Point(91, 114);
            this.Group.Name = "Group";
            this.Group.Size = new System.Drawing.Size(100, 22);
            this.Group.TabIndex = 14;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(24, 115);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Group";
            // 
            // btnQD
            // 
            this.btnQD.BackColor = System.Drawing.Color.Transparent;
            this.btnQD.Image = global::QueryDesigner.Properties.Resources._1303882176_search_16;
            this.btnQD.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnQD.Location = new System.Drawing.Point(197, 117);
            this.btnQD.Name = "btnQD";
            this.btnQD.Size = new System.Drawing.Size(16, 16);
            this.btnQD.TabIndex = 47;
            this.btnQD.TabStop = false;
            this.btnQD.Click += new System.EventHandler(this.btnQD_Click);
            // 
            // frmQDADD
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(912, 599);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.panelTab);
            this.Controls.Add(this.panelControl);
            this.Controls.Add(this.toolStrip1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "frmQDADD";
            this.Text = "Query Designer Address";
            this.Load += new System.EventHandler(this.frmQDADD_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.panelControl.ResumeLayout(false);
            this.panelControl.PerformLayout();
            this.panelTab.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabTable.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvFrom)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bsFROMCODE)).EndInit();
            this.tabField.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvField)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.bsField)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnQD)).EndInit();
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
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton btnTransferIn;
        private System.Windows.Forms.ToolStripButton btnTransferOut;
        private System.Windows.Forms.Panel panelControl;
        private System.Windows.Forms.Panel panelTab;
        private System.Windows.Forms.Button btnSelectTable;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabTable;
        private System.Windows.Forms.TabPage tabField;
        private System.Windows.Forms.Button btnSelectView;
        private System.Windows.Forms.BindingSource bsFROMCODE;
        private System.Windows.Forms.BindingSource bsField;
        private System.Windows.Forms.TextBox txtLookup;
        private System.Windows.Forms.Label lbLookup;
        private System.Windows.Forms.Button btnRelation;
        private System.Windows.Forms.CheckBox ckbUse;
        private System.Windows.Forms.TextBox txtModule;
        private System.Windows.Forms.Label label1;
        private Janus.Windows.GridEX.GridEX dgvField;
        private Janus.Windows.GridEX.GridEX dgvFrom;
        private System.Windows.Forms.TextBox txtConectEx;
        private System.Windows.Forms.ToolStripLabel lbErr;
        private System.Windows.Forms.Button btnXML;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem dgvAddRow;
        private System.Windows.Forms.TextBox Group;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox btnQD;
    }
}