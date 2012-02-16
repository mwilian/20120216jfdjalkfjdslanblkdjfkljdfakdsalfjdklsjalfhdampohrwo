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
		internal Janus.Windows.EditControls.UIGroupBox UiGroupBox1;
		internal Janus.Windows.ExplorerBar.ExplorerBar ExplorerBar1;
		internal Janus.Windows.UI.Tab.UITab UiTab1;
		internal System.Windows.Forms.ImageList ImageList1;
		internal Janus.Windows.ExplorerBar.ExplorerBar ExplorerBar2;
		internal Janus.Windows.ExplorerBar.ExplorerBarContainerControl ExplorerBarContainerControl4;
		internal Janus.Windows.EditControls.UIGroupBox UiGroupBox2;
		internal Janus.Windows.GridEX.EditControls.EditBox txtConditionName;
		internal Janus.Windows.GridEX.EditControls.EditBox txtValue2;
		internal Janus.Windows.GridEX.EditControls.EditBox txtValue1;
		internal Janus.Windows.EditControls.UIComboBox cboFields;
		internal System.Windows.Forms.Label lblFields;
		internal System.Windows.Forms.Label lblValue1;
		internal System.Windows.Forms.Label lblValue2;
		internal System.Windows.Forms.Label lblName;
		internal Janus.Windows.EditControls.UIComboBox cboCondition;
		internal System.Windows.Forms.Label lblCondition;
		internal Janus.Windows.EditControls.UIButton btnMoveDown;
		internal Janus.Windows.EditControls.UIButton btnMoveUp;
		internal Janus.Windows.EditControls.UIButton btnDelete;
		internal Janus.Windows.EditControls.UIButton btnNew;
		internal Janus.Windows.EditControls.UIButton btnCancel;
		internal Janus.Windows.EditControls.UIButton btnOK;
		internal Janus.Windows.GridEX.GridEX jsgConditions;
		internal Janus.Windows.ExplorerBar.ExplorerBarContainerControl excConditionName;
		internal Janus.Windows.ExplorerBar.ExplorerBarContainerControl excAppearance;
		internal Janus.Windows.ExplorerBar.ExplorerBarContainerControl excConditionCriteria;
		internal Janus.Windows.EditControls.UICheckBox chkStrikeout;
		internal Janus.Windows.EditControls.UICheckBox chkUnderline;
		internal Janus.Windows.EditControls.UICheckBox chkItalic;
		internal Janus.Windows.EditControls.UICheckBox chkBold;
		internal Janus.Windows.EditControls.UIColorButton btnBackColor;
		internal System.Windows.Forms.Label lblBackColor;
		internal System.Windows.Forms.Label lblForeColor;
		internal Janus.Windows.EditControls.UIColorButton btnForeColor;
		internal Janus.Windows.UI.Tab.UITabPage fontPage;
		internal Janus.Windows.UI.Tab.UITabPage colorsPage;
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            Janus.Windows.GridEX.GridEXLayout jsgConditions_DesignTimeLayout = new Janus.Windows.GridEX.GridEXLayout();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmFormatConditions));
            Janus.Windows.ExplorerBar.ExplorerBarGroup explorerBarGroup1 = new Janus.Windows.ExplorerBar.ExplorerBarGroup();
            Janus.Windows.ExplorerBar.ExplorerBarGroup explorerBarGroup2 = new Janus.Windows.ExplorerBar.ExplorerBarGroup();
            Janus.Windows.ExplorerBar.ExplorerBarGroup explorerBarGroup3 = new Janus.Windows.ExplorerBar.ExplorerBarGroup();
            Janus.Windows.ExplorerBar.ExplorerBarGroup explorerBarGroup4 = new Janus.Windows.ExplorerBar.ExplorerBarGroup();
            this.ExplorerBarContainerControl4 = new Janus.Windows.ExplorerBar.ExplorerBarContainerControl();
            this.jsgConditions = new Janus.Windows.GridEX.GridEX();
            this.ImageList1 = new System.Windows.Forms.ImageList(this.components);
            this.excConditionName = new Janus.Windows.ExplorerBar.ExplorerBarContainerControl();
            this.txtConditionName = new Janus.Windows.GridEX.EditControls.EditBox();
            this.lblName = new System.Windows.Forms.Label();
            this.excConditionCriteria = new Janus.Windows.ExplorerBar.ExplorerBarContainerControl();
            this.lblFields = new System.Windows.Forms.Label();
            this.txtValue2 = new Janus.Windows.GridEX.EditControls.EditBox();
            this.txtValue1 = new Janus.Windows.GridEX.EditControls.EditBox();
            this.cboCondition = new Janus.Windows.EditControls.UIComboBox();
            this.lblCondition = new System.Windows.Forms.Label();
            this.lblValue1 = new System.Windows.Forms.Label();
            this.lblValue2 = new System.Windows.Forms.Label();
            this.cboFields = new Janus.Windows.EditControls.UIComboBox();
            this.excAppearance = new Janus.Windows.ExplorerBar.ExplorerBarContainerControl();
            this.UiTab1 = new Janus.Windows.UI.Tab.UITab();
            this.fontPage = new Janus.Windows.UI.Tab.UITabPage();
            this.UiGroupBox2 = new Janus.Windows.EditControls.UIGroupBox();
            this.chkStrikeout = new Janus.Windows.EditControls.UICheckBox();
            this.chkUnderline = new Janus.Windows.EditControls.UICheckBox();
            this.chkItalic = new Janus.Windows.EditControls.UICheckBox();
            this.chkBold = new Janus.Windows.EditControls.UICheckBox();
            this.colorsPage = new Janus.Windows.UI.Tab.UITabPage();
            this.lblBackColor = new System.Windows.Forms.Label();
            this.btnBackColor = new Janus.Windows.EditControls.UIColorButton();
            this.lblForeColor = new System.Windows.Forms.Label();
            this.btnForeColor = new Janus.Windows.EditControls.UIColorButton();
            this.UiGroupBox1 = new Janus.Windows.EditControls.UIGroupBox();
            this.ExplorerBar2 = new Janus.Windows.ExplorerBar.ExplorerBar();
            this.ExplorerBar1 = new Janus.Windows.ExplorerBar.ExplorerBar();
            this.btnMoveDown = new Janus.Windows.EditControls.UIButton();
            this.btnMoveUp = new Janus.Windows.EditControls.UIButton();
            this.btnDelete = new Janus.Windows.EditControls.UIButton();
            this.btnNew = new Janus.Windows.EditControls.UIButton();
            this.btnCancel = new Janus.Windows.EditControls.UIButton();
            this.btnOK = new Janus.Windows.EditControls.UIButton();
            this.officeFormAdorner1 = new Janus.Windows.Ribbon.OfficeFormAdorner(this.components);
            this.ExplorerBarContainerControl4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.jsgConditions)).BeginInit();
            this.excConditionName.SuspendLayout();
            this.excConditionCriteria.SuspendLayout();
            this.excAppearance.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.UiTab1)).BeginInit();
            this.UiTab1.SuspendLayout();
            this.fontPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.UiGroupBox2)).BeginInit();
            this.UiGroupBox2.SuspendLayout();
            this.colorsPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.UiGroupBox1)).BeginInit();
            this.UiGroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ExplorerBar2)).BeginInit();
            this.ExplorerBar2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ExplorerBar1)).BeginInit();
            this.ExplorerBar1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.officeFormAdorner1)).BeginInit();
            this.SuspendLayout();
            // 
            // ExplorerBarContainerControl4
            // 
            this.ExplorerBarContainerControl4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.ExplorerBarContainerControl4.Controls.Add(this.jsgConditions);
            this.ExplorerBarContainerControl4.Location = new System.Drawing.Point(8, 27);
            this.ExplorerBarContainerControl4.Name = "ExplorerBarContainerControl4";
            this.ExplorerBarContainerControl4.Size = new System.Drawing.Size(252, 277);
            this.ExplorerBarContainerControl4.TabIndex = 1;
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
            this.jsgConditions.Location = new System.Drawing.Point(0, 0);
            this.jsgConditions.Name = "jsgConditions";
            this.jsgConditions.ScrollBarWidth = 17;
            this.jsgConditions.Size = new System.Drawing.Size(252, 277);
            this.jsgConditions.TabIndex = 0;
            this.jsgConditions.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            this.jsgConditions.UpdatingRecord += new System.ComponentModel.CancelEventHandler(this.jsgConditions_UpdatingRecord);
            this.jsgConditions.SelectionChanged += new System.EventHandler(this.jsgConditions_SelectionChanged);
            this.jsgConditions.CurrentCellChanging += new Janus.Windows.GridEX.CurrentCellChangingEventHandler(this.jsgConditions_CurrentCellChanging);
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
            this.excConditionName.Size = new System.Drawing.Size(280, 27);
            this.excConditionName.TabIndex = 1;
            // 
            // txtConditionName
            // 
            this.txtConditionName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtConditionName.ControlStyle.ButtonAppearance = Janus.Windows.GridEX.ButtonAppearance.Regular;
            this.txtConditionName.Location = new System.Drawing.Point(104, 4);
            this.txtConditionName.Name = "txtConditionName";
            this.txtConditionName.Size = new System.Drawing.Size(170, 21);
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
            this.excConditionCriteria.Size = new System.Drawing.Size(280, 101);
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
            this.txtValue2.Size = new System.Drawing.Size(200, 21);
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
            this.txtValue1.Size = new System.Drawing.Size(200, 21);
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
            this.cboCondition.Size = new System.Drawing.Size(200, 21);
            this.cboCondition.TabIndex = 1;
            this.cboCondition.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
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
            this.cboFields.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
            this.cboFields.Location = new System.Drawing.Point(72, 4);
            this.cboFields.Name = "cboFields";
            this.cboFields.Size = new System.Drawing.Size(200, 21);
            this.cboFields.TabIndex = 0;
            this.cboFields.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // excAppearance
            // 
            this.excAppearance.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.excAppearance.Controls.Add(this.UiTab1);
            this.excAppearance.Location = new System.Drawing.Point(8, 189);
            this.excAppearance.Name = "excAppearance";
            this.excAppearance.Size = new System.Drawing.Size(280, 115);
            this.excAppearance.TabIndex = 3;
            // 
            // UiTab1
            // 
            this.UiTab1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.UiTab1.FocusOnClick = false;
            this.UiTab1.ImageList = this.ImageList1;
            this.UiTab1.Location = new System.Drawing.Point(0, 0);
            this.UiTab1.Name = "UiTab1";
            this.UiTab1.Size = new System.Drawing.Size(280, 115);
            this.UiTab1.TabIndex = 0;
            this.UiTab1.TabPages.AddRange(new Janus.Windows.UI.Tab.UITabPage[] {
            this.fontPage,
            this.colorsPage});
            this.UiTab1.TabsStateStyles.SelectedFormatStyle.FontBold = Janus.Windows.UI.TriState.True;
            this.UiTab1.TabStripFormatStyle.BackgroundGradientMode = Janus.Windows.UI.BackgroundGradientMode.Solid;
            this.UiTab1.VisualStyle = Janus.Windows.UI.Tab.TabVisualStyle.Office2003;
            // 
            // fontPage
            // 
            this.fontPage.Controls.Add(this.UiGroupBox2);
            this.fontPage.ImageIndex = 0;
            this.fontPage.Location = new System.Drawing.Point(1, 23);
            this.fontPage.Name = "fontPage";
            this.fontPage.Size = new System.Drawing.Size(278, 91);
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
            this.UiGroupBox2.Text = "Font Style";
            this.UiGroupBox2.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
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
            this.chkStrikeout.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
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
            this.chkUnderline.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
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
            this.chkItalic.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
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
            this.chkBold.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // colorsPage
            // 
            this.colorsPage.Controls.Add(this.lblBackColor);
            this.colorsPage.Controls.Add(this.btnBackColor);
            this.colorsPage.Controls.Add(this.lblForeColor);
            this.colorsPage.Controls.Add(this.btnForeColor);
            this.colorsPage.ImageIndex = 3;
            this.colorsPage.Location = new System.Drawing.Point(1, 23);
            this.colorsPage.Name = "colorsPage";
            this.colorsPage.Size = new System.Drawing.Size(280, 127);
            this.colorsPage.TabStop = true;
            this.colorsPage.Text = "Colors";
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
            // btnBackColor
            // 
            // 
            // 
            // 
            this.btnBackColor.ColorPicker.AutomaticButtonText = "None";
            this.btnBackColor.ColorPicker.AutomaticColor = System.Drawing.Color.Empty;
            this.btnBackColor.ColorPicker.BorderStyle = Janus.Windows.UI.BorderStyle.None;
            this.btnBackColor.ColorPicker.Location = new System.Drawing.Point(0, 0);
            this.btnBackColor.ColorPicker.Name = "";
            this.btnBackColor.ColorPicker.Size = new System.Drawing.Size(100, 100);
            this.btnBackColor.ColorPicker.TabIndex = 0;
            this.btnBackColor.ImageHorizontalAlignment = Janus.Windows.EditControls.ImageHorizontalAlignment.Near;
            this.btnBackColor.ImageSize = new System.Drawing.Size(25, 15);
            this.btnBackColor.Location = new System.Drawing.Point(88, 40);
            this.btnBackColor.Name = "btnBackColor";
            this.btnBackColor.Size = new System.Drawing.Size(128, 23);
            this.btnBackColor.TabIndex = 0;
            this.btnBackColor.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnBackColor.WordWrap = false;
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
            // btnForeColor
            // 
            // 
            // 
            // 
            this.btnForeColor.ColorPicker.AutomaticButtonText = "None";
            this.btnForeColor.ColorPicker.AutomaticColor = System.Drawing.Color.Empty;
            this.btnForeColor.ColorPicker.BorderStyle = Janus.Windows.UI.BorderStyle.None;
            this.btnForeColor.ColorPicker.Location = new System.Drawing.Point(0, 0);
            this.btnForeColor.ColorPicker.Name = "";
            this.btnForeColor.ColorPicker.Size = new System.Drawing.Size(100, 100);
            this.btnForeColor.ColorPicker.TabIndex = 0;
            this.btnForeColor.ImageHorizontalAlignment = Janus.Windows.EditControls.ImageHorizontalAlignment.Near;
            this.btnForeColor.ImageSize = new System.Drawing.Size(25, 15);
            this.btnForeColor.Location = new System.Drawing.Point(88, 8);
            this.btnForeColor.Name = "btnForeColor";
            this.btnForeColor.Size = new System.Drawing.Size(128, 23);
            this.btnForeColor.TabIndex = 0;
            this.btnForeColor.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnForeColor.WordWrap = false;
            // 
            // UiGroupBox1
            // 
            this.UiGroupBox1.BackgroundStyle = Janus.Windows.EditControls.BackgroundStyle.Panel;
            this.UiGroupBox1.Controls.Add(this.ExplorerBar2);
            this.UiGroupBox1.Controls.Add(this.ExplorerBar1);
            this.UiGroupBox1.Controls.Add(this.btnMoveDown);
            this.UiGroupBox1.Controls.Add(this.btnMoveUp);
            this.UiGroupBox1.Controls.Add(this.btnDelete);
            this.UiGroupBox1.Controls.Add(this.btnNew);
            this.UiGroupBox1.Controls.Add(this.btnCancel);
            this.UiGroupBox1.Controls.Add(this.btnOK);
            this.UiGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.UiGroupBox1.FrameStyle = Janus.Windows.EditControls.FrameStyle.None;
            this.UiGroupBox1.Location = new System.Drawing.Point(0, 0);
            this.UiGroupBox1.Name = "UiGroupBox1";
            this.UiGroupBox1.Size = new System.Drawing.Size(576, 383);
            this.UiGroupBox1.TabIndex = 0;
            this.UiGroupBox1.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // ExplorerBar2
            // 
            this.ExplorerBar2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.ExplorerBar2.BackgroundThemeStyle = Janus.Windows.ExplorerBar.BackgroundThemeStyle.Items;
            this.ExplorerBar2.Controls.Add(this.ExplorerBarContainerControl4);
            explorerBarGroup1.Container = true;
            explorerBarGroup1.ContainerControl = this.ExplorerBarContainerControl4;
            explorerBarGroup1.ContainerHeight = 278;
            explorerBarGroup1.Expandable = false;
            explorerBarGroup1.Key = "Group1";
            explorerBarGroup1.SpecialGroup = true;
            explorerBarGroup1.Text = "Format Conditions";
            this.ExplorerBar2.Groups.AddRange(new Janus.Windows.ExplorerBar.ExplorerBarGroup[] {
            explorerBarGroup1});
            this.ExplorerBar2.GroupSeparation = 4;
            this.ExplorerBar2.Location = new System.Drawing.Point(4, 32);
            this.ExplorerBar2.Name = "ExplorerBar2";
            this.ExplorerBar2.Size = new System.Drawing.Size(268, 312);
            this.ExplorerBar2.TabIndex = 8;
            this.ExplorerBar2.Text = "ExplorerBar2";
            this.ExplorerBar2.VisualStyle = Janus.Windows.ExplorerBar.VisualStyle.Office2003;
            // 
            // ExplorerBar1
            // 
            this.ExplorerBar1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.ExplorerBar1.BackgroundFormatStyle.BackgroundThemeStyle = Janus.Windows.ExplorerBar.BackgroundThemeStyle.Items;
            this.ExplorerBar1.BackgroundThemeStyle = Janus.Windows.ExplorerBar.BackgroundThemeStyle.Items;
            this.ExplorerBar1.Controls.Add(this.excConditionName);
            this.ExplorerBar1.Controls.Add(this.excConditionCriteria);
            this.ExplorerBar1.Controls.Add(this.excAppearance);
            explorerBarGroup2.Container = true;
            explorerBarGroup2.ContainerControl = this.excConditionName;
            explorerBarGroup2.ContainerHeight = 28;
            explorerBarGroup2.Key = "Group1";
            explorerBarGroup2.ShowGroupCaption = false;
            explorerBarGroup2.Text = "New Group";
            explorerBarGroup3.Container = true;
            explorerBarGroup3.ContainerControl = this.excConditionCriteria;
            explorerBarGroup3.ContainerHeight = 102;
            explorerBarGroup3.Expandable = false;
            explorerBarGroup3.ImageIndex = 2;
            explorerBarGroup3.Key = "Group2";
            explorerBarGroup3.Text = "Condition";
            explorerBarGroup4.Container = true;
            explorerBarGroup4.ContainerControl = this.excAppearance;
            explorerBarGroup4.ContainerHeight = 116;
            explorerBarGroup4.Expandable = false;
            explorerBarGroup4.ImageIndex = 1;
            explorerBarGroup4.Key = "Group3";
            explorerBarGroup4.Text = "Appearence";
            this.ExplorerBar1.Groups.AddRange(new Janus.Windows.ExplorerBar.ExplorerBarGroup[] {
            explorerBarGroup2,
            explorerBarGroup3,
            explorerBarGroup4});
            this.ExplorerBar1.GroupSeparation = 6;
            this.ExplorerBar1.ImageList = this.ImageList1;
            this.ExplorerBar1.Location = new System.Drawing.Point(276, 32);
            this.ExplorerBar1.Name = "ExplorerBar1";
            this.ExplorerBar1.Size = new System.Drawing.Size(296, 312);
            this.ExplorerBar1.TabIndex = 7;
            this.ExplorerBar1.Text = "ExplorerBar1";
            this.ExplorerBar1.VisualStyle = Janus.Windows.ExplorerBar.VisualStyle.Office2003;
            // 
            // btnMoveDown
            // 
            this.btnMoveDown.Enabled = false;
            this.btnMoveDown.Location = new System.Drawing.Point(256, 4);
            this.btnMoveDown.Name = "btnMoveDown";
            this.btnMoveDown.Size = new System.Drawing.Size(80, 24);
            this.btnMoveDown.TabIndex = 6;
            this.btnMoveDown.Text = "Move Down";
            this.btnMoveDown.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnMoveDown.Click += new System.EventHandler(this.btnMoveDown_Click);
            // 
            // btnMoveUp
            // 
            this.btnMoveUp.Enabled = false;
            this.btnMoveUp.Location = new System.Drawing.Point(172, 4);
            this.btnMoveUp.Name = "btnMoveUp";
            this.btnMoveUp.Size = new System.Drawing.Size(80, 24);
            this.btnMoveUp.TabIndex = 5;
            this.btnMoveUp.Text = "Move Up";
            this.btnMoveUp.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnMoveUp.Click += new System.EventHandler(this.btnMoveUp_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(88, 4);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(80, 24);
            this.btnDelete.TabIndex = 3;
            this.btnDelete.Text = "Delete";
            this.btnDelete.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnNew
            // 
            this.btnNew.Location = new System.Drawing.Point(4, 4);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(80, 24);
            this.btnNew.TabIndex = 2;
            this.btnNew.Text = "New";
            this.btnNew.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(492, 352);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(80, 24);
            this.btnCancel.TabIndex = 10;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(404, 352);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(80, 24);
            this.btnOK.TabIndex = 9;
            this.btnOK.Text = "OK";
            this.btnOK.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // officeFormAdorner1
            // 
            this.officeFormAdorner1.DocumentName = "Automatic Formatting";
            this.officeFormAdorner1.Form = this;
            this.officeFormAdorner1.Office2007CustomColor = System.Drawing.Color.Empty;
            // 
            // frmFormatConditions
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
            this.ClientSize = new System.Drawing.Size(576, 383);
            this.Controls.Add(this.UiGroupBox1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmFormatConditions";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Automatic Formatting";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.FormatsForm_Closing);
            this.ExplorerBarContainerControl4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.jsgConditions)).EndInit();
            this.excConditionName.ResumeLayout(false);
            this.excConditionName.PerformLayout();
            this.excConditionCriteria.ResumeLayout(false);
            this.excConditionCriteria.PerformLayout();
            this.excAppearance.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.UiTab1)).EndInit();
            this.UiTab1.ResumeLayout(false);
            this.fontPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.UiGroupBox2)).EndInit();
            this.UiGroupBox2.ResumeLayout(false);
            this.colorsPage.ResumeLayout(false);
            this.colorsPage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.UiGroupBox1)).EndInit();
            this.UiGroupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ExplorerBar2)).EndInit();
            this.ExplorerBar2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ExplorerBar1)).EndInit();
            this.ExplorerBar1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.officeFormAdorner1)).EndInit();
            this.ResumeLayout(false);

		}

        private Janus.Windows.Ribbon.OfficeFormAdorner officeFormAdorner1;

	}

} //end of root namespace