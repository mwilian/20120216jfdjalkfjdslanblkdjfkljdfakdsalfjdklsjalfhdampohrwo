namespace QueryDesigner
{
    partial class FrmSystem
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
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabConnect = new System.Windows.Forms.TabPage();
            this.dgvList = new System.Windows.Forms.DataGridView();
            this.DEFAULT = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.KEY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Type = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.CONTENT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ContentEx = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TEST = new System.Windows.Forms.DataGridViewLinkColumn();
            this.BUILD = new System.Windows.Forms.DataGridViewButtonColumn();
            this.iTEMBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.bindingConfig = new System.Windows.Forms.BindingSource(this.components);
            this.qDConfig = new QueryDesigner.QDConfig();
            this.tabAP = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txtForce = new System.Windows.Forms.TextBox();
            this.sYSBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.txtBack = new System.Windows.Forms.TextBox();
            this.ckb2007 = new System.Windows.Forms.CheckBox();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnForce = new System.Windows.Forms.Button();
            this.panelBack = new System.Windows.Forms.Panel();
            this.panelForce = new System.Windows.Forms.Panel();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnFont = new System.Windows.Forms.Button();
            this.txtFont = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.PictureBox();
            this.label4 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtRPT = new System.Windows.Forms.TextBox();
            this.dIRBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.ddlAP = new System.Windows.Forms.ComboBox();
            this.dTBBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.ddlQD = new System.Windows.Forms.ComboBox();
            this.txtTMP = new System.Windows.Forms.TextBox();
            this.APITEMBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.QDITEMBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.panel1.SuspendLayout();
            this.tabControl.SuspendLayout();
            this.tabConnect.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iTEMBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingConfig)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.qDConfig)).BeginInit();
            this.tabAP.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sYSBindingSource)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.button1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.button2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dIRBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dTBBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.APITEMBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.QDITEMBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnOK);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 325);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(871, 36);
            this.panel1.TabIndex = 5;
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.Location = new System.Drawing.Point(712, 6);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(793, 6);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // tabControl
            // 
            this.tabControl.Appearance = System.Windows.Forms.TabAppearance.Buttons;
            this.tabControl.Controls.Add(this.tabConnect);
            this.tabControl.Controls.Add(this.tabAP);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(871, 325);
            this.tabControl.TabIndex = 6;
            this.tabControl.SelectedIndexChanged += new System.EventHandler(this.tabControl_SelectedIndexChanged);
            // 
            // tabConnect
            // 
            this.tabConnect.Controls.Add(this.dgvList);
            this.tabConnect.Location = new System.Drawing.Point(4, 25);
            this.tabConnect.Name = "tabConnect";
            this.tabConnect.Padding = new System.Windows.Forms.Padding(3);
            this.tabConnect.Size = new System.Drawing.Size(863, 296);
            this.tabConnect.TabIndex = 0;
            this.tabConnect.Text = "Connection";
            this.tabConnect.UseVisualStyleBackColor = true;
            // 
            // dgvList
            // 
            this.dgvList.AutoGenerateColumns = false;
            this.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DEFAULT,
            this.KEY,
            this.Type,
            this.CONTENT,
            this.ContentEx,
            this.TEST,
            this.BUILD});
            this.dgvList.DataSource = this.iTEMBindingSource;
            this.dgvList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvList.Location = new System.Drawing.Point(3, 3);
            this.dgvList.Name = "dgvList";
            this.dgvList.Size = new System.Drawing.Size(857, 290);
            this.dgvList.TabIndex = 0;
            this.dgvList.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvList_CellValueChanged);
            this.dgvList.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvList_CellClick);
            // 
            // DEFAULT
            // 
            this.DEFAULT.DataPropertyName = "DEFAULT";
            this.DEFAULT.HeaderText = "Default";
            this.DEFAULT.Name = "DEFAULT";
            this.DEFAULT.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.DEFAULT.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.DEFAULT.Width = 50;
            // 
            // KEY
            // 
            this.KEY.DataPropertyName = "KEY";
            this.KEY.HeaderText = "Key";
            this.KEY.Name = "KEY";
            // 
            // Type
            // 
            this.Type.DataPropertyName = "TYPE";
            this.Type.HeaderText = "Type";
            this.Type.Items.AddRange(new object[] {
            "QD",
            "AP"});
            this.Type.Name = "Type";
            this.Type.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.Type.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // CONTENT
            // 
            this.CONTENT.DataPropertyName = "CONTENT";
            this.CONTENT.HeaderText = "CONTENT";
            this.CONTENT.Name = "CONTENT";
            this.CONTENT.ReadOnly = true;
            this.CONTENT.Visible = false;
            // 
            // ContentEx
            // 
            this.ContentEx.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ContentEx.DataPropertyName = "CONTENTEX";
            this.ContentEx.HeaderText = "Connect String";
            this.ContentEx.Name = "ContentEx";
            // 
            // TEST
            // 
            this.TEST.DataPropertyName = "TEST";
            this.TEST.HeaderText = "Test";
            this.TEST.Name = "TEST";
            this.TEST.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.TEST.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // BUILD
            // 
            this.BUILD.DataPropertyName = "BUTTON";
            this.BUILD.HeaderText = "Build";
            this.BUILD.Name = "BUILD";
            this.BUILD.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.BUILD.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // iTEMBindingSource
            // 
            this.iTEMBindingSource.DataMember = "ITEM";
            this.iTEMBindingSource.DataSource = this.bindingConfig;
            // 
            // bindingConfig
            // 
            this.bindingConfig.DataSource = this.qDConfig;
            this.bindingConfig.Position = 0;
            // 
            // qDConfig
            // 
            this.qDConfig.DataSetName = "QDConfig";
            this.qDConfig.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // tabAP
            // 
            this.tabAP.Controls.Add(this.groupBox2);
            this.tabAP.Controls.Add(this.groupBox1);
            this.tabAP.Location = new System.Drawing.Point(4, 25);
            this.tabAP.Name = "tabAP";
            this.tabAP.Padding = new System.Windows.Forms.Padding(3);
            this.tabAP.Size = new System.Drawing.Size(863, 296);
            this.tabAP.TabIndex = 1;
            this.tabAP.Text = "Application";
            this.tabAP.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txtForce);
            this.groupBox2.Controls.Add(this.txtBack);
            this.groupBox2.Controls.Add(this.ckb2007);
            this.groupBox2.Controls.Add(this.btnBack);
            this.groupBox2.Controls.Add(this.btnForce);
            this.groupBox2.Controls.Add(this.panelBack);
            this.groupBox2.Controls.Add(this.panelForce);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.btnFont);
            this.groupBox2.Controls.Add(this.txtFont);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox2.Location = new System.Drawing.Point(3, 86);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(857, 107);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "System";
            // 
            // txtForce
            // 
            this.txtForce.DataBindings.Add(new System.Windows.Forms.Binding("Tag", this.sYSBindingSource, "BACKCOLOR", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtForce.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sYSBindingSource, "BACKCOLOR", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtForce.Location = new System.Drawing.Point(251, 51);
            this.txtForce.Name = "txtForce";
            this.txtForce.Size = new System.Drawing.Size(13, 20);
            this.txtForce.TabIndex = 14;
            this.txtForce.Visible = false;
            this.txtForce.TextChanged += new System.EventHandler(this.txtForce_TextChanged);
            // 
            // sYSBindingSource
            // 
            this.sYSBindingSource.DataMember = "SYS";
            this.sYSBindingSource.DataSource = this.bindingConfig;
            // 
            // txtBack
            // 
            this.txtBack.DataBindings.Add(new System.Windows.Forms.Binding("Tag", this.sYSBindingSource, "FORCECOLOR", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtBack.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sYSBindingSource, "FORCECOLOR", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtBack.Location = new System.Drawing.Point(251, 80);
            this.txtBack.Name = "txtBack";
            this.txtBack.Size = new System.Drawing.Size(13, 20);
            this.txtBack.TabIndex = 13;
            this.txtBack.Visible = false;
            // 
            // ckb2007
            // 
            this.ckb2007.AutoSize = true;
            this.ckb2007.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.sYSBindingSource, "USE2007", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.ckb2007.Location = new System.Drawing.Point(506, 22);
            this.ckb2007.Name = "ckb2007";
            this.ckb2007.Size = new System.Drawing.Size(101, 17);
            this.ckb2007.TabIndex = 12;
            this.ckb2007.Text = "Use Excel 2007";
            this.ckb2007.UseVisualStyleBackColor = true;
            // 
            // btnBack
            // 
            this.btnBack.Location = new System.Drawing.Point(170, 78);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(75, 23);
            this.btnBack.TabIndex = 11;
            this.btnBack.Text = "Default";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // btnForce
            // 
            this.btnForce.Location = new System.Drawing.Point(170, 49);
            this.btnForce.Name = "btnForce";
            this.btnForce.Size = new System.Drawing.Size(75, 23);
            this.btnForce.TabIndex = 10;
            this.btnForce.Text = "Default";
            this.btnForce.UseVisualStyleBackColor = true;
            this.btnForce.Click += new System.EventHandler(this.btnForce_Click);
            // 
            // panelBack
            // 
            this.panelBack.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelBack.Location = new System.Drawing.Point(89, 78);
            this.panelBack.Name = "panelBack";
            this.panelBack.Size = new System.Drawing.Size(75, 23);
            this.panelBack.TabIndex = 9;
            this.panelBack.Click += new System.EventHandler(this.panelBack_Click);
            // 
            // panelForce
            // 
            this.panelForce.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.panelForce.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelForce.Location = new System.Drawing.Point(89, 49);
            this.panelForce.Name = "panelForce";
            this.panelForce.Size = new System.Drawing.Size(75, 23);
            this.panelForce.TabIndex = 8;
            this.panelForce.Click += new System.EventHandler(this.panelForce_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(18, 83);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(59, 13);
            this.label7.TabIndex = 5;
            this.label7.Text = "Back Color";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(18, 54);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(61, 13);
            this.label6.TabIndex = 4;
            this.label6.Text = "Force Color";
            // 
            // btnFont
            // 
            this.btnFont.Location = new System.Drawing.Point(308, 21);
            this.btnFont.Name = "btnFont";
            this.btnFont.Size = new System.Drawing.Size(75, 23);
            this.btnFont.TabIndex = 3;
            this.btnFont.Text = "Choose";
            this.btnFont.UseVisualStyleBackColor = true;
            this.btnFont.Click += new System.EventHandler(this.btnFont_Click);
            // 
            // txtFont
            // 
            this.txtFont.DataBindings.Add(new System.Windows.Forms.Binding("Tag", this.sYSBindingSource, "FONT", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtFont.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sYSBindingSource, "FONT", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtFont.Location = new System.Drawing.Point(89, 23);
            this.txtFont.Name = "txtFont";
            this.txtFont.Size = new System.Drawing.Size(213, 20);
            this.txtFont.TabIndex = 2;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(18, 26);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 13);
            this.label5.TabIndex = 1;
            this.label5.Text = "System Font";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtRPT);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.ddlAP);
            this.groupBox1.Controls.Add(this.ddlQD);
            this.groupBox1.Controls.Add(this.txtTMP);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(857, 83);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Connect";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Transparent;
            this.button1.Image = global::QueryDesigner.Properties.Resources._1303882176_search_16;
            this.button1.Location = new System.Drawing.Point(680, 20);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(16, 16);
            this.button1.TabIndex = 49;
            this.button1.TabStop = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(263, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(39, 13);
            this.label4.TabIndex = 9;
            this.label4.Text = "Report";
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.Transparent;
            this.button2.Image = global::QueryDesigner.Properties.Resources._1303882176_search_16;
            this.button2.Location = new System.Drawing.Point(680, 47);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(16, 16);
            this.button2.TabIndex = 46;
            this.button2.TabStop = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Report Source";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(263, 21);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Template";
            // 
            // txtRPT
            // 
            this.txtRPT.DataBindings.Add(new System.Windows.Forms.Binding("Tag", this.dIRBindingSource, "RPT", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtRPT.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.dIRBindingSource, "RPT", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtRPT.Location = new System.Drawing.Point(322, 45);
            this.txtRPT.Name = "txtRPT";
            this.txtRPT.Size = new System.Drawing.Size(352, 20);
            this.txtRPT.TabIndex = 4;
            // 
            // dIRBindingSource
            // 
            this.dIRBindingSource.DataMember = "DIR";
            this.dIRBindingSource.DataSource = this.bindingConfig;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Query Source";
            // 
            // ddlAP
            // 
            this.ddlAP.DataBindings.Add(new System.Windows.Forms.Binding("SelectedItem", this.dTBBindingSource, "AP", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.ddlAP.DataBindings.Add(new System.Windows.Forms.Binding("SelectedValue", this.dTBBindingSource, "AP", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.ddlAP.DataBindings.Add(new System.Windows.Forms.Binding("Tag", this.dTBBindingSource, "AP", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.ddlAP.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.dTBBindingSource, "AP", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.ddlAP.DataSource = this.APITEMBindingSource;
            this.ddlAP.DisplayMember = "KEY";
            this.ddlAP.FormattingEnabled = true;
            this.ddlAP.Location = new System.Drawing.Point(89, 45);
            this.ddlAP.Name = "ddlAP";
            this.ddlAP.Size = new System.Drawing.Size(121, 21);
            this.ddlAP.TabIndex = 6;
            this.ddlAP.ValueMember = "KEY";
            // 
            // dTBBindingSource
            // 
            this.dTBBindingSource.DataMember = "DTB";
            this.dTBBindingSource.DataSource = this.bindingConfig;
            // 
            // ddlQD
            // 
            this.ddlQD.DataBindings.Add(new System.Windows.Forms.Binding("SelectedItem", this.dTBBindingSource, "QD", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.ddlQD.DataBindings.Add(new System.Windows.Forms.Binding("SelectedValue", this.dTBBindingSource, "QD", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.ddlQD.DataBindings.Add(new System.Windows.Forms.Binding("Tag", this.dTBBindingSource, "QD", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.ddlQD.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.dTBBindingSource, "QD", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.ddlQD.DataSource = this.QDITEMBindingSource;
            this.ddlQD.DisplayMember = "KEY";
            this.ddlQD.FormattingEnabled = true;
            this.ddlQD.Location = new System.Drawing.Point(89, 18);
            this.ddlQD.Name = "ddlQD";
            this.ddlQD.Size = new System.Drawing.Size(121, 21);
            this.ddlQD.TabIndex = 5;
            this.ddlQD.ValueMember = "KEY";
            // 
            // txtTMP
            // 
            this.txtTMP.DataBindings.Add(new System.Windows.Forms.Binding("Tag", this.dIRBindingSource, "TMP", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtTMP.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.dIRBindingSource, "TMP", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtTMP.Location = new System.Drawing.Point(322, 18);
            this.txtTMP.Name = "txtTMP";
            this.txtTMP.Size = new System.Drawing.Size(352, 20);
            this.txtTMP.TabIndex = 3;
            // 
            // APITEMBindingSource
            // 
            this.APITEMBindingSource.DataMember = "ITEM";
            this.APITEMBindingSource.DataSource = this.bindingConfig;
            this.APITEMBindingSource.Filter = "TYPE=\'AP\'";
            // 
            // QDITEMBindingSource
            // 
            this.QDITEMBindingSource.DataMember = "ITEM";
            this.QDITEMBindingSource.DataSource = this.bindingConfig;
            this.QDITEMBindingSource.Filter = "TYPE=\'QD\'";
            // 
            // FrmSystem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(871, 361);
            this.Controls.Add(this.tabControl);
            this.Controls.Add(this.panel1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.MaximumSize = new System.Drawing.Size(9999, 9999);
            this.Name = "FrmSystem";
            this.Text = "Connection System";
            this.Load += new System.EventHandler(this.FrmSystem_Load);
            this.panel1.ResumeLayout(false);
            this.tabControl.ResumeLayout(false);
            this.tabConnect.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iTEMBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingConfig)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.qDConfig)).EndInit();
            this.tabAP.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sYSBindingSource)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.button1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.button2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dIRBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dTBBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.APITEMBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.QDITEMBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabConnect;
        private System.Windows.Forms.DataGridView dgvList;
        private System.Windows.Forms.TabPage tabAP;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.Button btnForce;
        private System.Windows.Forms.Panel panelBack;
        private System.Windows.Forms.Panel panelForce;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnFont;
        private System.Windows.Forms.TextBox txtFont;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtRPT;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox ddlAP;
        private System.Windows.Forms.ComboBox ddlQD;
        private System.Windows.Forms.TextBox txtTMP;
        private System.Windows.Forms.PictureBox button1;
        private System.Windows.Forms.PictureBox button2;
        private System.Windows.Forms.CheckBox ckb2007;
        private System.Windows.Forms.BindingSource iTEMBindingSource;
        private System.Windows.Forms.BindingSource bindingConfig;
        private QDConfig qDConfig;
        private System.Windows.Forms.BindingSource dTBBindingSource;
        private System.Windows.Forms.BindingSource sYSBindingSource;
        private System.Windows.Forms.BindingSource dIRBindingSource;
        private System.Windows.Forms.TextBox txtForce;
        private System.Windows.Forms.TextBox txtBack;
        private System.Windows.Forms.DataGridViewCheckBoxColumn DEFAULT;
        private System.Windows.Forms.DataGridViewTextBoxColumn KEY;
        private System.Windows.Forms.DataGridViewComboBoxColumn Type;
        private System.Windows.Forms.DataGridViewTextBoxColumn CONTENT;
        private System.Windows.Forms.DataGridViewTextBoxColumn ContentEx;
        private System.Windows.Forms.DataGridViewLinkColumn TEST;
        private System.Windows.Forms.DataGridViewButtonColumn BUILD;
        private System.Windows.Forms.BindingSource APITEMBindingSource;
        private System.Windows.Forms.BindingSource QDITEMBindingSource;
    }
}