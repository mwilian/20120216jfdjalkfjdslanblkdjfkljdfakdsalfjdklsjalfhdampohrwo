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
	public partial class ConditionalGroupControl : System.Windows.Forms.UserControl
	{


		//UserControl overrides dispose to clean up the component list.
		[System.Diagnostics.DebuggerNonUserCode()]
		protected override void Dispose(bool disposing)
		{
			if (disposing && components != null)
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		//Required by the Windows Form Designer
		private System.ComponentModel.IContainer components;

		//NOTE: The following procedure is required by the Windows Form Designer
		//It can be modified using the Windows Form Designer.  
		//Do not modify it using the code editor.
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConditionalGroupControl));
            this.grbBackground = new Janus.Windows.EditControls.UIGroupBox();
            this.chkShowWhenEmpty = new Janus.Windows.EditControls.UICheckBox();
            this.lblGroupRowCaption = new System.Windows.Forms.Label();
            this.txtGroupRowCaption = new Janus.Windows.GridEX.EditControls.EditBox();
            this.btnMoveDown = new Janus.Windows.EditControls.UIButton();
            this.ImageList1 = new System.Windows.Forms.ImageList(this.components);
            this.btnMoveUp = new Janus.Windows.EditControls.UIButton();
            this.btnRemoveGroupRow = new Janus.Windows.EditControls.UIButton();
            this.btnNewGroupRow = new Janus.Windows.EditControls.UIButton();
            this.FilterEditor1 = new Janus.Windows.FilterEditor.FilterEditor();
            this.grdGroupRows = new Janus.Windows.GridEX.GridEX();
            this.txtHeaderCaption = new Janus.Windows.GridEX.EditControls.EditBox();
            this.lblGroupCaption = new System.Windows.Forms.Label();
            this.lblName = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.txtName = new Janus.Windows.GridEX.EditControls.EditBox();
            this.btnNewCustomGroup = new Janus.Windows.EditControls.UIButton();
            this.cboSelectCustomGroup = new Janus.Windows.EditControls.UIComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.grbBackground)).BeginInit();
            this.grbBackground.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdGroupRows)).BeginInit();
            this.SuspendLayout();
            // 
            // grbBackground
            // 
            this.grbBackground.BackgroundStyle = Janus.Windows.EditControls.BackgroundStyle.Panel;
            this.grbBackground.Controls.Add(this.chkShowWhenEmpty);
            this.grbBackground.Controls.Add(this.lblGroupRowCaption);
            this.grbBackground.Controls.Add(this.txtGroupRowCaption);
            this.grbBackground.Controls.Add(this.btnMoveDown);
            this.grbBackground.Controls.Add(this.btnMoveUp);
            this.grbBackground.Controls.Add(this.btnRemoveGroupRow);
            this.grbBackground.Controls.Add(this.btnNewGroupRow);
            this.grbBackground.Controls.Add(this.FilterEditor1);
            this.grbBackground.Controls.Add(this.grdGroupRows);
            this.grbBackground.Controls.Add(this.txtHeaderCaption);
            this.grbBackground.Controls.Add(this.lblGroupCaption);
            this.grbBackground.Controls.Add(this.lblName);
            this.grbBackground.Controls.Add(this.Label1);
            this.grbBackground.Controls.Add(this.txtName);
            this.grbBackground.Controls.Add(this.btnNewCustomGroup);
            this.grbBackground.Controls.Add(this.cboSelectCustomGroup);
            this.grbBackground.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grbBackground.FrameStyle = Janus.Windows.EditControls.FrameStyle.None;
            this.grbBackground.Location = new System.Drawing.Point(0, 0);
            this.grbBackground.Name = "grbBackground";
            this.grbBackground.Size = new System.Drawing.Size(718, 372);
            this.grbBackground.TabIndex = 8;
            this.grbBackground.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // chkShowWhenEmpty
            // 
            this.chkShowWhenEmpty.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkShowWhenEmpty.BackColor = System.Drawing.Color.Transparent;
            this.chkShowWhenEmpty.Enabled = false;
            this.chkShowWhenEmpty.Location = new System.Drawing.Point(218, 342);
            this.chkShowWhenEmpty.Name = "chkShowWhenEmpty";
            this.chkShowWhenEmpty.Size = new System.Drawing.Size(104, 17);
            this.chkShowWhenEmpty.TabIndex = 24;
            this.chkShowWhenEmpty.Text = "Show when empty";
            this.chkShowWhenEmpty.VisualStyle = Janus.Windows.UI.VisualStyle.VS2005;
            this.chkShowWhenEmpty.CheckedChanged += new System.EventHandler(this.chkShowWhenEmpty_CheckedChanged);
            // 
            // lblGroupRowCaption
            // 
            this.lblGroupRowCaption.AutoSize = true;
            this.lblGroupRowCaption.BackColor = System.Drawing.Color.Transparent;
            this.lblGroupRowCaption.Enabled = false;
            this.lblGroupRowCaption.Location = new System.Drawing.Point(215, 104);
            this.lblGroupRowCaption.Name = "lblGroupRowCaption";
            this.lblGroupRowCaption.Size = new System.Drawing.Size(100, 13);
            this.lblGroupRowCaption.TabIndex = 23;
            this.lblGroupRowCaption.Text = "Group Row Caption";
            // 
            // txtGroupRowCaption
            // 
            this.txtGroupRowCaption.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtGroupRowCaption.Enabled = false;
            this.txtGroupRowCaption.Location = new System.Drawing.Point(321, 100);
            this.txtGroupRowCaption.Name = "txtGroupRowCaption";
            this.txtGroupRowCaption.Size = new System.Drawing.Size(384, 20);
            this.txtGroupRowCaption.TabIndex = 22;
            this.txtGroupRowCaption.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            this.txtGroupRowCaption.TextChanged += new System.EventHandler(this.txtGroupRowCaption_TextChanged);
            // 
            // btnMoveDown
            // 
            this.btnMoveDown.Enabled = false;
            this.btnMoveDown.ImageIndex = 1;
            this.btnMoveDown.ImageList = this.ImageList1;
            this.btnMoveDown.Location = new System.Drawing.Point(317, 70);
            this.btnMoveDown.Name = "btnMoveDown";
            this.btnMoveDown.Size = new System.Drawing.Size(93, 23);
            this.btnMoveDown.TabIndex = 20;
            this.btnMoveDown.Text = "Move Down";
            this.btnMoveDown.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnMoveDown.Click += new System.EventHandler(this.btnMoveDown_Click);
            // 
            // ImageList1
            // 
            this.ImageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("ImageList1.ImageStream")));
            this.ImageList1.TransparentColor = System.Drawing.Color.Magenta;
            this.ImageList1.Images.SetKeyName(0, "Conditionalgroups.bmp");
            this.ImageList1.Images.SetKeyName(1, "MoveDown.bmp");
            this.ImageList1.Images.SetKeyName(2, "MoveUp.bmp");
            this.ImageList1.Images.SetKeyName(3, "NewConditional.bmp");
            this.ImageList1.Images.SetKeyName(4, "RemoveConditional.bmp");
            // 
            // btnMoveUp
            // 
            this.btnMoveUp.Enabled = false;
            this.btnMoveUp.ImageIndex = 2;
            this.btnMoveUp.ImageList = this.ImageList1;
            this.btnMoveUp.Location = new System.Drawing.Point(216, 70);
            this.btnMoveUp.Name = "btnMoveUp";
            this.btnMoveUp.Size = new System.Drawing.Size(93, 23);
            this.btnMoveUp.TabIndex = 19;
            this.btnMoveUp.Text = "Move Up";
            this.btnMoveUp.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnMoveUp.Click += new System.EventHandler(this.btnMoveUp_Click);
            // 
            // btnRemoveGroupRow
            // 
            this.btnRemoveGroupRow.Enabled = false;
            this.btnRemoveGroupRow.ImageIndex = 4;
            this.btnRemoveGroupRow.ImageList = this.ImageList1;
            this.btnRemoveGroupRow.Location = new System.Drawing.Point(115, 70);
            this.btnRemoveGroupRow.Name = "btnRemoveGroupRow";
            this.btnRemoveGroupRow.Size = new System.Drawing.Size(93, 23);
            this.btnRemoveGroupRow.TabIndex = 18;
            this.btnRemoveGroupRow.Text = "Remove";
            this.btnRemoveGroupRow.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnRemoveGroupRow.Click += new System.EventHandler(this.btnRemoveGroupRow_Click);
            // 
            // btnNewGroupRow
            // 
            this.btnNewGroupRow.Enabled = false;
            this.btnNewGroupRow.ImageIndex = 3;
            this.btnNewGroupRow.ImageList = this.ImageList1;
            this.btnNewGroupRow.Location = new System.Drawing.Point(14, 70);
            this.btnNewGroupRow.Name = "btnNewGroupRow";
            this.btnNewGroupRow.Size = new System.Drawing.Size(93, 23);
            this.btnNewGroupRow.TabIndex = 17;
            this.btnNewGroupRow.Text = "New";
            this.btnNewGroupRow.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnNewGroupRow.Click += new System.EventHandler(this.btnNewGroupRow_Click);
            // 
            // FilterEditor1
            // 
            this.FilterEditor1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.FilterEditor1.BackColor = System.Drawing.Color.Transparent;
            this.FilterEditor1.Enabled = false;
            this.FilterEditor1.InnerAreaStyle = Janus.Windows.UI.Dock.PanelInnerAreaStyle.UseFormatStyle;
            this.FilterEditor1.Location = new System.Drawing.Point(218, 128);
            this.FilterEditor1.MinSize = new System.Drawing.Size(0, 0);
            this.FilterEditor1.Name = "FilterEditor1";
            this.FilterEditor1.ScrollMode = Janus.Windows.UI.Dock.ScrollMode.Both;
            this.FilterEditor1.ScrollStep = 15;
            this.FilterEditor1.Size = new System.Drawing.Size(487, 202);
            this.FilterEditor1.FilterConditionChanged += new System.EventHandler(this.FilterEditor1_FilterConditionChanged);
            // 
            // grdGroupRows
            // 
            this.grdGroupRows.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.False;
            this.grdGroupRows.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.grdGroupRows.ColumnHeaders = Janus.Windows.GridEX.InheritableBoolean.False;
            this.grdGroupRows.Enabled = false;
            this.grdGroupRows.GridLines = Janus.Windows.GridEX.GridLines.None;
            this.grdGroupRows.GroupByBoxVisible = false;
            this.grdGroupRows.HideSelection = Janus.Windows.GridEX.HideSelection.Highlight;
            this.grdGroupRows.Location = new System.Drawing.Point(14, 100);
            this.grdGroupRows.Name = "grdGroupRows";
            this.grdGroupRows.SaveSettings = false;
            this.grdGroupRows.Size = new System.Drawing.Size(188, 262);
            this.grdGroupRows.TabIndex = 16;
            this.grdGroupRows.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            this.grdGroupRows.SelectionChanged += new System.EventHandler(this.grdGroupRows_SelectionChanged);
            // 
            // txtHeaderCaption
            // 
            this.txtHeaderCaption.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtHeaderCaption.Location = new System.Drawing.Point(408, 39);
            this.txtHeaderCaption.Name = "txtHeaderCaption";
            this.txtHeaderCaption.Size = new System.Drawing.Size(297, 20);
            this.txtHeaderCaption.TabIndex = 14;
            this.txtHeaderCaption.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            this.txtHeaderCaption.TextChanged += new System.EventHandler(this.txtHeaderCaption_TextChanged);
            // 
            // lblGroupCaption
            // 
            this.lblGroupCaption.AutoSize = true;
            this.lblGroupCaption.BackColor = System.Drawing.Color.Transparent;
            this.lblGroupCaption.Location = new System.Drawing.Point(289, 43);
            this.lblGroupCaption.Name = "lblGroupCaption";
            this.lblGroupCaption.Size = new System.Drawing.Size(113, 13);
            this.lblGroupCaption.TabIndex = 13;
            this.lblGroupCaption.Text = "Group Header Caption";
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.BackColor = System.Drawing.Color.Transparent;
            this.lblName.Enabled = false;
            this.lblName.Location = new System.Drawing.Point(14, 43);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(35, 13);
            this.lblName.TabIndex = 12;
            this.lblName.Text = "Name";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.BackColor = System.Drawing.Color.Transparent;
            this.Label1.Location = new System.Drawing.Point(14, 13);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(148, 13);
            this.Label1.TabIndex = 11;
            this.Label1.Text = "Choose existing custom group";
            // 
            // txtName
            // 
            this.txtName.Enabled = false;
            this.txtName.Location = new System.Drawing.Point(55, 39);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(228, 20);
            this.txtName.TabIndex = 10;
            this.txtName.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            this.txtName.TextChanged += new System.EventHandler(this.txtName_TextChanged);
            // 
            // btnNewCustomGroup
            // 
            this.btnNewCustomGroup.ImageIndex = 0;
            this.btnNewCustomGroup.ImageList = this.ImageList1;
            this.btnNewCustomGroup.Location = new System.Drawing.Point(354, 8);
            this.btnNewCustomGroup.Name = "btnNewCustomGroup";
            this.btnNewCustomGroup.Size = new System.Drawing.Size(75, 23);
            this.btnNewCustomGroup.TabIndex = 9;
            this.btnNewCustomGroup.Text = "New";
            this.btnNewCustomGroup.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnNewCustomGroup.Click += new System.EventHandler(this.btnNewCustomGroup_Click);
            // 
            // cboSelectCustomGroup
            // 
            this.cboSelectCustomGroup.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
            this.cboSelectCustomGroup.Location = new System.Drawing.Point(168, 9);
            this.cboSelectCustomGroup.Name = "cboSelectCustomGroup";
            this.cboSelectCustomGroup.Size = new System.Drawing.Size(176, 20);
            this.cboSelectCustomGroup.TabIndex = 8;
            this.cboSelectCustomGroup.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.cboSelectCustomGroup.SelectedItemChanged += new System.EventHandler(this.cboSelectCustomGroup_SelectedItemChanged);
            // 
            // ConditionalGroupControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.grbBackground);
            this.Name = "ConditionalGroupControl";
            this.Size = new System.Drawing.Size(718, 372);
            ((System.ComponentModel.ISupportInitialize)(this.grbBackground)).EndInit();
            this.grbBackground.ResumeLayout(false);
            this.grbBackground.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdGroupRows)).EndInit();
            this.ResumeLayout(false);

		}

		public ConditionalGroupControl()
		{

			// This call is required by the Windows Form Designer.
			InitializeComponent();

			// Add any initialization after the InitializeComponent() call.

		}
		internal Janus.Windows.EditControls.UIGroupBox grbBackground;
		internal Janus.Windows.GridEX.EditControls.EditBox txtHeaderCaption;
		internal System.Windows.Forms.Label lblGroupCaption;
		internal System.Windows.Forms.Label lblName;
		internal System.Windows.Forms.Label Label1;
		internal Janus.Windows.GridEX.EditControls.EditBox txtName;
		internal Janus.Windows.EditControls.UIButton btnNewCustomGroup;
		internal Janus.Windows.EditControls.UIComboBox cboSelectCustomGroup;
		internal Janus.Windows.EditControls.UIButton btnMoveDown;
		internal Janus.Windows.EditControls.UIButton btnMoveUp;
		internal Janus.Windows.EditControls.UIButton btnRemoveGroupRow;
		internal Janus.Windows.EditControls.UIButton btnNewGroupRow;
		internal Janus.Windows.FilterEditor.FilterEditor FilterEditor1;
		internal Janus.Windows.GridEX.GridEX grdGroupRows;
		internal Janus.Windows.EditControls.UICheckBox chkShowWhenEmpty;
		internal System.Windows.Forms.Label lblGroupRowCaption;
		internal Janus.Windows.GridEX.EditControls.EditBox txtGroupRowCaption;
		internal System.Windows.Forms.ImageList ImageList1;
	}

} //end of root namespace