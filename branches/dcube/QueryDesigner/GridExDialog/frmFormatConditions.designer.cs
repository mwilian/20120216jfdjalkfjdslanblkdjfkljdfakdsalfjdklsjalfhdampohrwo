using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;

using Janus.Windows.GridEX;

namespace QueryDesigner
{
    public partial class frmFormatConditions : System.Windows.Forms.Form
    {


        //Form overrides dispose to clean up the component list.
        internal frmFormatConditions()
        {
            InitializeComponent();
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        //Required by the Windows Form Designer
        private System.ComponentModel.IContainer components;

        //NOTE: The following procedure is required by the Windows Form Designer
        //It can be modified using the Windows Form Designer.  
        //Do not modify it using the code editor.
        internal GroupBox UiGroupBox1;
        internal Panel ExplorerBar1;
        internal TabControl UiTab1;
        internal System.Windows.Forms.ImageList ImageList1;
        internal GroupBox UiGroupBox2;
        internal Janus.Windows.GridEX.EditControls.EditBox txtConditionName;
        internal Janus.Windows.GridEX.EditControls.EditBox txtValue2;
        internal Janus.Windows.GridEX.EditControls.EditBox txtValue1;
        internal ComboBox cboFields;
        internal System.Windows.Forms.Label lblFields;
        internal System.Windows.Forms.Label lblValue1;
        internal System.Windows.Forms.Label lblValue2;
        internal System.Windows.Forms.Label lblName;
        internal ComboBox cboCondition;
        internal System.Windows.Forms.Label lblCondition;
        internal Button btnMoveDown;
        internal Button btnMoveUp;
        internal Button btnDelete;
        internal Button btnNew;
        internal Button btnCancel;
        internal Button btnOK;
        internal Janus.Windows.GridEX.GridEX jsgConditions;
        internal Panel excConditionName;
        internal Panel excConditionCriteria;
        internal CheckBox chkStrikeout;
        internal CheckBox chkUnderline;
        internal CheckBox chkItalic;
        internal CheckBox chkBold;
        internal System.Windows.Forms.Label lblBackColor;
        internal System.Windows.Forms.Label lblForeColor;
        internal TabPage fontPage;
        internal TabPage colorsPage;
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            Janus.Windows.GridEX.GridEXLayout jsgConditions_DesignTimeLayout = new Janus.Windows.GridEX.GridEXLayout();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmFormatConditions));
            this.explorerBarGroup2 = new System.Windows.Forms.Panel();
            this.jsgConditions = new Janus.Windows.GridEX.GridEX();
            this.ImageList1 = new System.Windows.Forms.ImageList(this.components);
            this.excConditionName = new System.Windows.Forms.Panel();
            this.txtConditionName = new Janus.Windows.GridEX.EditControls.EditBox();
            this.lblName = new System.Windows.Forms.Label();
            this.excConditionCriteria = new System.Windows.Forms.Panel();
            this.lblFields = new System.Windows.Forms.Label();
            this.txtValue2 = new Janus.Windows.GridEX.EditControls.EditBox();
            this.txtValue1 = new Janus.Windows.GridEX.EditControls.EditBox();
            this.cboCondition = new System.Windows.Forms.ComboBox();
            this.lblCondition = new System.Windows.Forms.Label();
            this.lblValue1 = new System.Windows.Forms.Label();
            this.lblValue2 = new System.Windows.Forms.Label();
            this.cboFields = new System.Windows.Forms.ComboBox();
            this.UiTab1 = new System.Windows.Forms.TabControl();
            this.fontPage = new System.Windows.Forms.TabPage();
            this.UiGroupBox2 = new System.Windows.Forms.GroupBox();
            this.chkStrikeout = new System.Windows.Forms.CheckBox();
            this.chkUnderline = new System.Windows.Forms.CheckBox();
            this.chkItalic = new System.Windows.Forms.CheckBox();
            this.chkBold = new System.Windows.Forms.CheckBox();
            this.colorsPage = new System.Windows.Forms.TabPage();
            this.btnBackColor = new System.Windows.Forms.Button();
            this.btnForeColor = new System.Windows.Forms.Button();
            this.lblBackColor = new System.Windows.Forms.Label();
            this.lblForeColor = new System.Windows.Forms.Label();
            this.UiGroupBox1 = new System.Windows.Forms.GroupBox();
            this.ExplorerBar1 = new System.Windows.Forms.Panel();
            this.btnMoveDown = new System.Windows.Forms.Button();
            this.btnMoveUp = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.jsgConditions)).BeginInit();
            this.excConditionName.SuspendLayout();
            this.excConditionCriteria.SuspendLayout();
            this.UiTab1.SuspendLayout();
            this.fontPage.SuspendLayout();
            this.UiGroupBox2.SuspendLayout();
            this.colorsPage.SuspendLayout();
            this.UiGroupBox1.SuspendLayout();
            this.ExplorerBar1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // explorerBarGroup2
            // 
            this.explorerBarGroup2.Location = new System.Drawing.Point(0, 0);
            this.explorerBarGroup2.Name = "explorerBarGroup2";
            this.explorerBarGroup2.Size = new System.Drawing.Size(311, 45);
            this.explorerBarGroup2.TabIndex = 4;
            this.explorerBarGroup2.Text = "New Group";
            // 
            // jsgConditions
            // 
            this.jsgConditions.BorderStyle = Janus.Windows.GridEX.BorderStyle.None;
            this.jsgConditions.ColumnAutoResize = true;
            jsgConditions_DesignTimeLayout.LayoutString = resources.GetString("jsgConditions_DesignTimeLayout.LayoutString");
            this.jsgConditions.DesignTimeLayout = jsgConditions_DesignTimeLayout;
            this.jsgConditions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.jsgConditions.GroupByBoxVisible = false;
            this.jsgConditions.ImageList = this.ImageList1;
            this.jsgConditions.Location = new System.Drawing.Point(3, 17);
            this.jsgConditions.Name = "jsgConditions";
            this.jsgConditions.ScrollBarWidth = 17;
            this.jsgConditions.Size = new System.Drawing.Size(276, 364);
            this.jsgConditions.TabIndex = 0;
            this.jsgConditions.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            this.jsgConditions.CurrentCellChanging += new Janus.Windows.GridEX.CurrentCellChangingEventHandler(this.jsgConditions_CurrentCellChanging);
            this.jsgConditions.UpdatingRecord += new System.ComponentModel.CancelEventHandler(this.jsgConditions_UpdatingRecord);
            this.jsgConditions.SelectionChanged += new System.EventHandler(this.jsgConditions_SelectionChanged);
            this.jsgConditions.FormattingRow += new Janus.Windows.GridEX.RowLoadEventHandler(this.jsgConditions_FormattingRow);
            // 
            // ImageList1
            // 
            this.ImageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("ImageList1.ImageStream")));
            this.ImageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.ImageList1.Images.SetKeyName(0, "");
            this.ImageList1.Images.SetKeyName(1, "");
            this.ImageList1.Images.SetKeyName(2, "");
            this.ImageList1.Images.SetKeyName(3, "");
            this.ImageList1.Images.SetKeyName(4, "");
            this.ImageList1.Images.SetKeyName(5, "");
            // 
            // excConditionName
            // 
            this.excConditionName.Controls.Add(this.txtConditionName);
            this.excConditionName.Controls.Add(this.lblName);
            this.excConditionName.Location = new System.Drawing.Point(8, 7);
            this.excConditionName.Name = "excConditionName";
            this.excConditionName.Size = new System.Drawing.Size(295, 27);
            this.excConditionName.TabIndex = 1;
            // 
            // txtConditionName
            // 
            this.txtConditionName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtConditionName.ControlStyle.ButtonAppearance = Janus.Windows.GridEX.ButtonAppearance.Regular;
            this.txtConditionName.Location = new System.Drawing.Point(104, 4);
            this.txtConditionName.Name = "txtConditionName";
            this.txtConditionName.Size = new System.Drawing.Size(185, 21);
            this.txtConditionName.TabIndex = 0;
            this.txtConditionName.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near;
            this.txtConditionName.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            this.txtConditionName.TextChanged += new System.EventHandler(this.txtConditionName_TextChanged);
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.BackColor = System.Drawing.Color.Transparent;
            this.lblName.Location = new System.Drawing.Point(8, 8);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(85, 13);
            this.lblName.TabIndex = 4;
            this.lblName.Text = "Condition name:";
            // 
            // excConditionCriteria
            // 
            this.excConditionCriteria.Controls.Add(this.lblFields);
            this.excConditionCriteria.Controls.Add(this.txtValue2);
            this.excConditionCriteria.Controls.Add(this.txtValue1);
            this.excConditionCriteria.Controls.Add(this.cboCondition);
            this.excConditionCriteria.Controls.Add(this.lblCondition);
            this.excConditionCriteria.Controls.Add(this.lblValue1);
            this.excConditionCriteria.Controls.Add(this.lblValue2);
            this.excConditionCriteria.Controls.Add(this.cboFields);
            this.excConditionCriteria.Location = new System.Drawing.Point(8, 61);
            this.excConditionCriteria.Name = "excConditionCriteria";
            this.excConditionCriteria.Size = new System.Drawing.Size(295, 101);
            this.excConditionCriteria.TabIndex = 2;
            // 
            // lblFields
            // 
            this.lblFields.AutoSize = true;
            this.lblFields.BackColor = System.Drawing.Color.Transparent;
            this.lblFields.Location = new System.Drawing.Point(8, 7);
            this.lblFields.Name = "lblFields";
            this.lblFields.Size = new System.Drawing.Size(38, 13);
            this.lblFields.TabIndex = 4;
            this.lblFields.Text = "Fields:";
            // 
            // txtValue2
            // 
            this.txtValue2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtValue2.ControlStyle.ButtonAppearance = Janus.Windows.GridEX.ButtonAppearance.Regular;
            this.txtValue2.Location = new System.Drawing.Point(72, 76);
            this.txtValue2.Name = "txtValue2";
            this.txtValue2.Size = new System.Drawing.Size(215, 21);
            this.txtValue2.TabIndex = 3;
            this.txtValue2.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near;
            this.txtValue2.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            // 
            // txtValue1
            // 
            this.txtValue1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtValue1.ControlStyle.ButtonAppearance = Janus.Windows.GridEX.ButtonAppearance.Regular;
            this.txtValue1.Location = new System.Drawing.Point(72, 52);
            this.txtValue1.Name = "txtValue1";
            this.txtValue1.Size = new System.Drawing.Size(215, 21);
            this.txtValue1.TabIndex = 2;
            this.txtValue1.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near;
            this.txtValue1.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            // 
            // cboCondition
            // 
            this.cboCondition.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cboCondition.Location = new System.Drawing.Point(72, 28);
            this.cboCondition.Name = "cboCondition";
            this.cboCondition.Size = new System.Drawing.Size(215, 21);
            this.cboCondition.TabIndex = 1;
            this.cboCondition.SelectedIndexChanged += new System.EventHandler(this.cboCondition_SelectedIndexChanged);
            // 
            // lblCondition
            // 
            this.lblCondition.AutoSize = true;
            this.lblCondition.BackColor = System.Drawing.Color.Transparent;
            this.lblCondition.Location = new System.Drawing.Point(8, 32);
            this.lblCondition.Name = "lblCondition";
            this.lblCondition.Size = new System.Drawing.Size(56, 13);
            this.lblCondition.TabIndex = 4;
            this.lblCondition.Text = "Condition:";
            // 
            // lblValue1
            // 
            this.lblValue1.AutoSize = true;
            this.lblValue1.BackColor = System.Drawing.Color.Transparent;
            this.lblValue1.Location = new System.Drawing.Point(8, 56);
            this.lblValue1.Name = "lblValue1";
            this.lblValue1.Size = new System.Drawing.Size(46, 13);
            this.lblValue1.TabIndex = 4;
            this.lblValue1.Text = "Value 1:";
            // 
            // lblValue2
            // 
            this.lblValue2.AutoSize = true;
            this.lblValue2.BackColor = System.Drawing.Color.Transparent;
            this.lblValue2.Location = new System.Drawing.Point(8, 80);
            this.lblValue2.Name = "lblValue2";
            this.lblValue2.Size = new System.Drawing.Size(46, 13);
            this.lblValue2.TabIndex = 4;
            this.lblValue2.Text = "Value 2:";
            // 
            // cboFields
            // 
            this.cboFields.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cboFields.Location = new System.Drawing.Point(72, 4);
            this.cboFields.Name = "cboFields";
            this.cboFields.Size = new System.Drawing.Size(215, 21);
            this.cboFields.TabIndex = 0;
            // 
            // UiTab1
            // 
            this.UiTab1.Controls.Add(this.fontPage);
            this.UiTab1.Controls.Add(this.colorsPage);
            this.UiTab1.ImageList = this.ImageList1;
            this.UiTab1.Location = new System.Drawing.Point(8, 164);
            this.UiTab1.Name = "UiTab1";
            this.UiTab1.SelectedIndex = 0;
            this.UiTab1.Size = new System.Drawing.Size(295, 196);
            this.UiTab1.TabIndex = 0;
            // 
            // fontPage
            // 
            this.fontPage.Controls.Add(this.UiGroupBox2);
            this.fontPage.ImageIndex = 0;
            this.fontPage.Location = new System.Drawing.Point(4, 23);
            this.fontPage.Name = "fontPage";
            this.fontPage.Size = new System.Drawing.Size(287, 169);
            this.fontPage.TabIndex = 0;
            this.fontPage.TabStop = true;
            this.fontPage.Text = "Font";
            // 
            // UiGroupBox2
            // 
            this.UiGroupBox2.BackColor = System.Drawing.Color.Transparent;
            this.UiGroupBox2.Controls.Add(this.chkStrikeout);
            this.UiGroupBox2.Controls.Add(this.chkUnderline);
            this.UiGroupBox2.Controls.Add(this.chkItalic);
            this.UiGroupBox2.Controls.Add(this.chkBold);
            this.UiGroupBox2.Location = new System.Drawing.Point(8, 8);
            this.UiGroupBox2.Name = "UiGroupBox2";
            this.UiGroupBox2.Size = new System.Drawing.Size(248, 40);
            this.UiGroupBox2.TabIndex = 0;
            this.UiGroupBox2.TabStop = false;
            this.UiGroupBox2.Text = "Font Style";
            // 
            // chkStrikeout
            // 
            this.chkStrikeout.BackColor = System.Drawing.Color.Transparent;
            this.chkStrikeout.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkStrikeout.Location = new System.Drawing.Point(176, 16);
            this.chkStrikeout.Name = "chkStrikeout";
            this.chkStrikeout.Size = new System.Drawing.Size(64, 16);
            this.chkStrikeout.TabIndex = 7;
            this.chkStrikeout.Text = "Strikeout";
            this.chkStrikeout.UseVisualStyleBackColor = false;
            // 
            // chkUnderline
            // 
            this.chkUnderline.BackColor = System.Drawing.Color.Transparent;
            this.chkUnderline.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkUnderline.Location = new System.Drawing.Point(104, 16);
            this.chkUnderline.Name = "chkUnderline";
            this.chkUnderline.Size = new System.Drawing.Size(72, 16);
            this.chkUnderline.TabIndex = 6;
            this.chkUnderline.Text = "Underline";
            this.chkUnderline.UseVisualStyleBackColor = false;
            // 
            // chkItalic
            // 
            this.chkItalic.BackColor = System.Drawing.Color.Transparent;
            this.chkItalic.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkItalic.Location = new System.Drawing.Point(56, 16);
            this.chkItalic.Name = "chkItalic";
            this.chkItalic.Size = new System.Drawing.Size(48, 16);
            this.chkItalic.TabIndex = 5;
            this.chkItalic.Text = "Italic";
            this.chkItalic.UseVisualStyleBackColor = false;
            // 
            // chkBold
            // 
            this.chkBold.BackColor = System.Drawing.Color.Transparent;
            this.chkBold.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkBold.Location = new System.Drawing.Point(8, 16);
            this.chkBold.Name = "chkBold";
            this.chkBold.Size = new System.Drawing.Size(48, 16);
            this.chkBold.TabIndex = 4;
            this.chkBold.Text = "Bold";
            this.chkBold.UseVisualStyleBackColor = false;
            // 
            // colorsPage
            // 
            this.colorsPage.Controls.Add(this.btnBackColor);
            this.colorsPage.Controls.Add(this.btnForeColor);
            this.colorsPage.Controls.Add(this.lblBackColor);
            this.colorsPage.Controls.Add(this.lblForeColor);
            this.colorsPage.ImageIndex = 3;
            this.colorsPage.Location = new System.Drawing.Point(4, 23);
            this.colorsPage.Name = "colorsPage";
            this.colorsPage.Size = new System.Drawing.Size(287, 508);
            this.colorsPage.TabIndex = 1;
            this.colorsPage.TabStop = true;
            this.colorsPage.Text = "Colors";
            // 
            // btnBackColor
            // 
            this.btnBackColor.Location = new System.Drawing.Point(88, 44);
            this.btnBackColor.Name = "btnBackColor";
            this.btnBackColor.Size = new System.Drawing.Size(128, 23);
            this.btnBackColor.TabIndex = 3;
            this.btnBackColor.Text = "(none)";
            this.btnBackColor.UseVisualStyleBackColor = true;
            this.btnBackColor.Click += new System.EventHandler(this.btnBackColor_Click);
            // 
            // btnForeColor
            // 
            this.btnForeColor.Location = new System.Drawing.Point(88, 12);
            this.btnForeColor.Name = "btnForeColor";
            this.btnForeColor.Size = new System.Drawing.Size(128, 23);
            this.btnForeColor.TabIndex = 2;
            this.btnForeColor.Text = "(none)";
            this.btnForeColor.UseVisualStyleBackColor = true;
            this.btnForeColor.Click += new System.EventHandler(this.btnForeColor_Click_1);
            // 
            // lblBackColor
            // 
            this.lblBackColor.AutoSize = true;
            this.lblBackColor.Location = new System.Drawing.Point(8, 44);
            this.lblBackColor.Name = "lblBackColor";
            this.lblBackColor.Size = new System.Drawing.Size(67, 13);
            this.lblBackColor.TabIndex = 1;
            this.lblBackColor.Text = "Background:";
            // 
            // lblForeColor
            // 
            this.lblForeColor.AutoSize = true;
            this.lblForeColor.Location = new System.Drawing.Point(8, 12);
            this.lblForeColor.Name = "lblForeColor";
            this.lblForeColor.Size = new System.Drawing.Size(33, 13);
            this.lblForeColor.TabIndex = 1;
            this.lblForeColor.Text = "Text:";
            // 
            // UiGroupBox1
            // 
            this.UiGroupBox1.Controls.Add(this.jsgConditions);
            this.UiGroupBox1.Controls.Add(this.ExplorerBar1);
            this.UiGroupBox1.Controls.Add(this.btnCancel);
            this.UiGroupBox1.Controls.Add(this.btnOK);
            this.UiGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.UiGroupBox1.Location = new System.Drawing.Point(0, 30);
            this.UiGroupBox1.Name = "UiGroupBox1";
            this.UiGroupBox1.Size = new System.Drawing.Size(593, 384);
            this.UiGroupBox1.TabIndex = 0;
            this.UiGroupBox1.TabStop = false;
            // 
            // ExplorerBar1
            // 
            this.ExplorerBar1.Controls.Add(this.UiTab1);
            this.ExplorerBar1.Controls.Add(this.excConditionName);
            this.ExplorerBar1.Controls.Add(this.excConditionCriteria);
            this.ExplorerBar1.Controls.Add(this.explorerBarGroup2);
            this.ExplorerBar1.Dock = System.Windows.Forms.DockStyle.Right;
            this.ExplorerBar1.Location = new System.Drawing.Point(279, 17);
            this.ExplorerBar1.Name = "ExplorerBar1";
            this.ExplorerBar1.Size = new System.Drawing.Size(311, 364);
            this.ExplorerBar1.TabIndex = 7;
            this.ExplorerBar1.Text = "ExplorerBar1";
            // 
            // btnMoveDown
            // 
            this.btnMoveDown.Enabled = false;
            this.btnMoveDown.Location = new System.Drawing.Point(257, 3);
            this.btnMoveDown.Name = "btnMoveDown";
            this.btnMoveDown.Size = new System.Drawing.Size(80, 24);
            this.btnMoveDown.TabIndex = 6;
            this.btnMoveDown.Text = "Move Down";
            this.btnMoveDown.Click += new System.EventHandler(this.btnMoveDown_Click);
            // 
            // btnMoveUp
            // 
            this.btnMoveUp.Enabled = false;
            this.btnMoveUp.Location = new System.Drawing.Point(171, 3);
            this.btnMoveUp.Name = "btnMoveUp";
            this.btnMoveUp.Size = new System.Drawing.Size(80, 24);
            this.btnMoveUp.TabIndex = 5;
            this.btnMoveUp.Text = "Move Up";
            this.btnMoveUp.Click += new System.EventHandler(this.btnMoveUp_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(85, 3);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(80, 24);
            this.btnDelete.TabIndex = 3;
            this.btnDelete.Text = "Delete";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnNew
            // 
            this.btnNew.Location = new System.Drawing.Point(3, 3);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(80, 24);
            this.btnNew.TabIndex = 2;
            this.btnNew.Text = "New";
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(887, 770);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(80, 24);
            this.btnCancel.TabIndex = 10;
            this.btnCancel.Text = "Cancel";
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(799, 770);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(80, 24);
            this.btnOK.TabIndex = 9;
            this.btnOK.Text = "OK";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnDelete);
            this.panel1.Controls.Add(this.btnNew);
            this.panel1.Controls.Add(this.btnMoveDown);
            this.panel1.Controls.Add(this.btnMoveUp);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(593, 30);
            this.panel1.TabIndex = 1;
            // 
            // frmFormatConditions
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
            this.ClientSize = new System.Drawing.Size(593, 414);
            this.Controls.Add(this.UiGroupBox1);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmFormatConditions";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Automatic Formatting";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.FormatsForm_Closing);
            ((System.ComponentModel.ISupportInitialize)(this.jsgConditions)).EndInit();
            this.excConditionName.ResumeLayout(false);
            this.excConditionName.PerformLayout();
            this.excConditionCriteria.ResumeLayout(false);
            this.excConditionCriteria.PerformLayout();
            this.UiTab1.ResumeLayout(false);
            this.fontPage.ResumeLayout(false);
            this.UiGroupBox2.ResumeLayout(false);
            this.colorsPage.ResumeLayout(false);
            this.colorsPage.PerformLayout();
            this.UiGroupBox1.ResumeLayout(false);
            this.ExplorerBar1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        private ColorDialog colorDialog1;
        private Button btnBackColor;
        private Button btnForeColor;
        private Panel explorerBarGroup2;
        private Panel panel1;

    }

} //end of root namespace