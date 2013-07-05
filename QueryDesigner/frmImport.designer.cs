using System.Windows.Forms;
namespace dCube
{
    partial class frmImport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmImport));
            this.btnBrowse = new Button();
            this.textBox1 = new TextBox();
            this.btnImport = new Button();
            this.ofdImport = new OpenFileDialog();
            this.btnPreview = new Button();
            this.imageList1 = new ImageList(this.components);
            this.label1 = new Label();
            this.ddlImport = new ComboBox();
            this.checkBox1 = new CheckBox();
            this.checkBox2 = new CheckBox();
            this.cboConvertor = new ComboBox();
            this.label2 = new Label();
            this.statusStrip1 = new StatusStrip();
            this.lbErr = new ToolStripStatusLabel();
            this.btnGroup = new Button();
            this.groupBox1 = new GroupBox();
            this.btnClear = new Button();
            this.btnAdd = new Button();
            this.tcMain = new TabControl();
            this.statusStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            this.btnBrowse.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.btnBrowse.Location = new System.Drawing.Point(925, 46);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 0;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((AnchorStyles)(((AnchorStyles.Top | AnchorStyles.Left)
                        | AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(6, 46);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(913, 48);
            this.textBox1.TabIndex = 1;
            // 
            // btnImport
            // 
            this.btnImport.Anchor = AnchorStyles.Top;
            this.btnImport.Enabled = false;
            this.btnImport.Location = new System.Drawing.Point(168, 17);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 2;
            this.btnImport.Text = "Import >>";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // ofdImport
            // 
            this.ofdImport.Filter = "Excel File (*.xls)|*.xls";
            // 
            // btnPreview
            // 
            this.btnPreview.Anchor = AnchorStyles.Top;
            this.btnPreview.Location = new System.Drawing.Point(6, 17);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(75, 23);
            this.btnPreview.TabIndex = 4;
            this.btnPreview.Text = "Preview >>";
            this.btnPreview.UseVisualStyleBackColor = true;
            this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "1307436862_alert.png");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 104);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Import Code";
            // 
            // ddlImport
            // 
            this.ddlImport.DisplayMember = "SCHEMA_ID";
            this.ddlImport.DropDownStyle = ComboBoxStyle.DropDownList;
            this.ddlImport.Location = new System.Drawing.Point(78, 100);
            this.ddlImport.Name = "ddlImport";
            this.ddlImport.Size = new System.Drawing.Size(179, 21);
            this.ddlImport.TabIndex = 7;
            this.ddlImport.ValueMember = "SCHEMA_ID";
            this.ddlImport.SelectedValueChanged += new System.EventHandler(this.dllImport_ValueChanged);
            // 
            // checkBox1
            // 
            this.checkBox1.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(518, 102);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(52, 17);
            this.checkBox1.TabIndex = 8;
            this.checkBox1.Text = "Insert";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(598, 102);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(61, 17);
            this.checkBox2.TabIndex = 9;
            this.checkBox2.Text = "Update";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // cboConvertor
            // 
            this.cboConvertor.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.cboConvertor.Items.AddRange(new object[] {
            "Unicode",
            "VNI",
            "TVCN3"});
            this.cboConvertor.Location = new System.Drawing.Point(762, 100);
            this.cboConvertor.Name = "cboConvertor";
            this.cboConvertor.Size = new System.Drawing.Size(157, 21);
            this.cboConvertor.TabIndex = 10;
            this.cboConvertor.Text = "None";
            // 
            // label2
            // 
            this.label2.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.label2.Location = new System.Drawing.Point(687, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Convert from";
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new ToolStripItem[] {
            this.lbErr});
            this.statusStrip1.Location = new System.Drawing.Point(0, 486);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1072, 22);
            this.statusStrip1.TabIndex = 12;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // lbErr
            // 
            this.lbErr.AutoSize = false;
            this.lbErr.ForeColor = System.Drawing.Color.BlueViolet;
            this.lbErr.Name = "lbErr";
            this.lbErr.Size = new System.Drawing.Size(200, 17);
            this.lbErr.Text = "...";
            this.lbErr.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lbErr.Click += new System.EventHandler(this.lbErr_Click);
            // 
            // btnGroup
            // 
            this.btnGroup.Anchor = AnchorStyles.Top;
            this.btnGroup.Enabled = false;
            this.btnGroup.Location = new System.Drawing.Point(87, 17);
            this.btnGroup.Name = "btnGroup";
            this.btnGroup.Size = new System.Drawing.Size(75, 23);
            this.btnGroup.TabIndex = 13;
            this.btnGroup.Text = "Group >>";
            this.btnGroup.UseVisualStyleBackColor = true;
            this.btnGroup.Click += new System.EventHandler(this.btnGroup_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnClear);
            this.groupBox1.Controls.Add(this.btnBrowse);
            this.groupBox1.Controls.Add(this.btnAdd);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnGroup);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.btnImport);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnPreview);
            this.groupBox1.Controls.Add(this.cboConvertor);
            this.groupBox1.Controls.Add(this.ddlImport);
            this.groupBox1.Controls.Add(this.checkBox2);
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Dock = DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1072, 129);
            this.groupBox1.TabIndex = 14;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "General";
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(344, 98);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 15;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(263, 98);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 14;
            this.btnAdd.Text = "Add >>";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // tcMain
            // 
            this.tcMain.Dock = DockStyle.Fill;
            this.tcMain.Location = new System.Drawing.Point(0, 129);
            this.tcMain.Name = "tcMain";
            this.tcMain.SelectedIndex = 0;
            this.tcMain.Size = new System.Drawing.Size(1072, 357);
            this.tcMain.TabIndex = 15;
            // 
            // frmImport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.ClientSize = new System.Drawing.Size(1072, 508);
            this.Controls.Add(this.tcMain);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.statusStrip1);
            this.Name = "frmImport";
            this.Text = "Import";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button btnBrowse;
        private TextBox textBox1;
        private Button btnImport;
        private OpenFileDialog ofdImport;
        private Button btnPreview;
        private Label label1;
        private ComboBox ddlImport;
        private CheckBox checkBox1;
        private CheckBox checkBox2;
        private ImageList imageList1;
        private ComboBox cboConvertor;
        private Label label2;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel lbErr;
        private Button btnGroup;
        private GroupBox groupBox1;
        private Button btnAdd;
        private Button btnClear;
        private TabControl tcMain;
    }
}

