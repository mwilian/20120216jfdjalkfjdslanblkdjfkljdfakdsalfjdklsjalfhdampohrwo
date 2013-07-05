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
	public partial class CompositeColumnsGroupcontrol : System.Windows.Forms.UserControl
	{


		//UserControl overrides dispose to clean up the component list.
		internal CompositeColumnsGroupcontrol()
		{
			InitializeComponent();
		}
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CompositeColumnsGroupcontrol));
            this.grbBackground = new Janus.Windows.EditControls.UIGroupBox();
            this.lblCompositeColumns = new System.Windows.Forms.Label();
            this.lblAvailableColumns = new System.Windows.Forms.Label();
            this.btnMoveDown = new Janus.Windows.EditControls.UIButton();
            this.ImageList1 = new System.Windows.Forms.ImageList(this.components);
            this.btnMoveUp = new Janus.Windows.EditControls.UIButton();
            this.btnRemove = new Janus.Windows.EditControls.UIButton();
            this.btnAdd = new Janus.Windows.EditControls.UIButton();
            this.grdCompositeColumns = new Janus.Windows.GridEX.GridEX();
            this.grdColumnList = new Janus.Windows.GridEX.GridEX();
            this.lblSelectTable = new System.Windows.Forms.Label();
            this.cboTable = new Janus.Windows.EditControls.UIComboBox();
            this.grbTable = new Janus.Windows.EditControls.UIGroupBox();
            ((System.ComponentModel.ISupportInitialize)(this.grbBackground)).BeginInit();
            this.grbBackground.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdCompositeColumns)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdColumnList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grbTable)).BeginInit();
            this.grbTable.SuspendLayout();
            this.SuspendLayout();
            // 
            // grbBackground
            // 
            this.grbBackground.BackgroundStyle = Janus.Windows.EditControls.BackgroundStyle.Panel;
            this.grbBackground.Controls.Add(this.lblCompositeColumns);
            this.grbBackground.Controls.Add(this.lblAvailableColumns);
            this.grbBackground.Controls.Add(this.btnMoveDown);
            this.grbBackground.Controls.Add(this.btnMoveUp);
            this.grbBackground.Controls.Add(this.btnRemove);
            this.grbBackground.Controls.Add(this.btnAdd);
            this.grbBackground.Controls.Add(this.grdCompositeColumns);
            this.grbBackground.Controls.Add(this.grdColumnList);
            this.grbBackground.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grbBackground.FrameStyle = Janus.Windows.EditControls.FrameStyle.None;
            this.grbBackground.Location = new System.Drawing.Point(0, 33);
            this.grbBackground.Name = "grbBackground";
            this.grbBackground.Size = new System.Drawing.Size(558, 290);
            this.grbBackground.TabIndex = 0;
            this.grbBackground.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // lblCompositeColumns
            // 
            this.lblCompositeColumns.AutoSize = true;
            this.lblCompositeColumns.BackColor = System.Drawing.Color.Transparent;
            this.lblCompositeColumns.Location = new System.Drawing.Point(301, 5);
            this.lblCompositeColumns.Name = "lblCompositeColumns";
            this.lblCompositeColumns.Size = new System.Drawing.Size(89, 13);
            this.lblCompositeColumns.TabIndex = 9;
            this.lblCompositeColumns.Text = "Columns to group";
            // 
            // lblAvailableColumns
            // 
            this.lblAvailableColumns.AutoSize = true;
            this.lblAvailableColumns.BackColor = System.Drawing.Color.Transparent;
            this.lblAvailableColumns.Location = new System.Drawing.Point(7, 5);
            this.lblAvailableColumns.Name = "lblAvailableColumns";
            this.lblAvailableColumns.Size = new System.Drawing.Size(93, 13);
            this.lblAvailableColumns.TabIndex = 8;
            this.lblAvailableColumns.Text = "Available Columns";
            // 
            // btnMoveDown
            // 
            this.btnMoveDown.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMoveDown.Enabled = false;
            this.btnMoveDown.ImageIndex = 2;
            this.btnMoveDown.ImageList = this.ImageList1;
            this.btnMoveDown.Location = new System.Drawing.Point(451, 254);
            this.btnMoveDown.Name = "btnMoveDown";
            this.btnMoveDown.Size = new System.Drawing.Size(92, 23);
            this.btnMoveDown.TabIndex = 6;
            this.btnMoveDown.Text = "Move Down";
            this.btnMoveDown.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnMoveDown.Click += new System.EventHandler(this.btnMoveDown_Click);
            // 
            // ImageList1
            // 
            this.ImageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("ImageList1.ImageStream")));
            this.ImageList1.TransparentColor = System.Drawing.Color.Magenta;
            this.ImageList1.Images.SetKeyName(0, "AddField.bmp");
            this.ImageList1.Images.SetKeyName(1, "RemoveField.bmp");
            this.ImageList1.Images.SetKeyName(2, "MoveDown.bmp");
            this.ImageList1.Images.SetKeyName(3, "MoveUp.bmp");
            // 
            // btnMoveUp
            // 
            this.btnMoveUp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMoveUp.Enabled = false;
            this.btnMoveUp.ImageIndex = 3;
            this.btnMoveUp.ImageList = this.ImageList1;
            this.btnMoveUp.Location = new System.Drawing.Point(353, 254);
            this.btnMoveUp.Name = "btnMoveUp";
            this.btnMoveUp.Size = new System.Drawing.Size(92, 23);
            this.btnMoveUp.TabIndex = 5;
            this.btnMoveUp.Text = "Move Up";
            this.btnMoveUp.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnMoveUp.Click += new System.EventHandler(this.btnMoveUp_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Enabled = false;
            this.btnRemove.ImageHorizontalAlignment = Janus.Windows.EditControls.ImageHorizontalAlignment.Near;
            this.btnRemove.ImageIndex = 1;
            this.btnRemove.ImageList = this.ImageList1;
            this.btnRemove.Location = new System.Drawing.Point(231, 54);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(67, 23);
            this.btnRemove.TabIndex = 4;
            this.btnRemove.Text = "Remove";
            this.btnRemove.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Enabled = false;
            this.btnAdd.ImageHorizontalAlignment = Janus.Windows.EditControls.ImageHorizontalAlignment.Far;
            this.btnAdd.ImageIndex = 0;
            this.btnAdd.ImageList = this.ImageList1;
            this.btnAdd.Location = new System.Drawing.Point(231, 25);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(67, 23);
            this.btnAdd.TabIndex = 3;
            this.btnAdd.Text = "Add";
            this.btnAdd.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // grdCompositeColumns
            // 
            this.grdCompositeColumns.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.False;
            this.grdCompositeColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.grdCompositeColumns.ColumnHeaders = Janus.Windows.GridEX.InheritableBoolean.False;
            this.grdCompositeColumns.GridLines = Janus.Windows.GridEX.GridLines.None;
            this.grdCompositeColumns.GroupByBoxVisible = false;
            this.grdCompositeColumns.HideSelection = Janus.Windows.GridEX.HideSelection.Highlight;
            this.grdCompositeColumns.Location = new System.Drawing.Point(304, 25);
            this.grdCompositeColumns.Name = "grdCompositeColumns";
            this.grdCompositeColumns.SaveSettings = false;
            this.grdCompositeColumns.Size = new System.Drawing.Size(239, 223);
            this.grdCompositeColumns.TabIndex = 2;
            this.grdCompositeColumns.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            this.grdCompositeColumns.SelectionChanged += new System.EventHandler(this.grdCompositeColumns_SelectionChanged);
            this.grdCompositeColumns.RowDoubleClick += new Janus.Windows.GridEX.RowActionEventHandler(this.grdCompositeColumns_RowDoubleClick);
            // 
            // grdColumnList
            // 
            this.grdColumnList.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.False;
            this.grdColumnList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.grdColumnList.ColumnHeaders = Janus.Windows.GridEX.InheritableBoolean.False;
            this.grdColumnList.GridLines = Janus.Windows.GridEX.GridLines.None;
            this.grdColumnList.GroupByBoxVisible = false;
            this.grdColumnList.HideSelection = Janus.Windows.GridEX.HideSelection.Highlight;
            this.grdColumnList.Location = new System.Drawing.Point(10, 25);
            this.grdColumnList.Name = "grdColumnList";
            this.grdColumnList.SaveSettings = false;
            this.grdColumnList.Size = new System.Drawing.Size(215, 223);
            this.grdColumnList.TabIndex = 0;
            this.grdColumnList.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            this.grdColumnList.SelectionChanged += new System.EventHandler(this.grdColumnList_SelectionChanged);
            this.grdColumnList.RowDoubleClick += new Janus.Windows.GridEX.RowActionEventHandler(this.grdColumnList_RowDoubleClick);
            // 
            // lblSelectTable
            // 
            this.lblSelectTable.AutoSize = true;
            this.lblSelectTable.BackColor = System.Drawing.Color.Transparent;
            this.lblSelectTable.Location = new System.Drawing.Point(9, 14);
            this.lblSelectTable.Name = "lblSelectTable";
            this.lblSelectTable.Size = new System.Drawing.Size(63, 13);
            this.lblSelectTable.TabIndex = 7;
            this.lblSelectTable.Text = "Select from:";
            // 
            // cboTable
            // 
            this.cboTable.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
            this.cboTable.Location = new System.Drawing.Point(78, 10);
            this.cboTable.Name = "cboTable";
            this.cboTable.Size = new System.Drawing.Size(187, 20);
            this.cboTable.TabIndex = 1;
            this.cboTable.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // grbTable
            // 
            this.grbTable.BackgroundStyle = Janus.Windows.EditControls.BackgroundStyle.Panel;
            this.grbTable.Controls.Add(this.cboTable);
            this.grbTable.Controls.Add(this.lblSelectTable);
            this.grbTable.Dock = System.Windows.Forms.DockStyle.Top;
            this.grbTable.FrameStyle = Janus.Windows.EditControls.FrameStyle.None;
            this.grbTable.Location = new System.Drawing.Point(0, 0);
            this.grbTable.Name = "grbTable";
            this.grbTable.Size = new System.Drawing.Size(558, 33);
            this.grbTable.TabIndex = 1;
            this.grbTable.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // CompositeColumnsGroupcontrol
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.grbBackground);
            this.Controls.Add(this.grbTable);
            this.Name = "CompositeColumnsGroupcontrol";
            this.Size = new System.Drawing.Size(558, 323);
            ((System.ComponentModel.ISupportInitialize)(this.grbBackground)).EndInit();
            this.grbBackground.ResumeLayout(false);
            this.grbBackground.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdCompositeColumns)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdColumnList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grbTable)).EndInit();
            this.grbTable.ResumeLayout(false);
            this.grbTable.PerformLayout();
            this.ResumeLayout(false);

		}
		internal Janus.Windows.EditControls.UIGroupBox grbBackground;
		internal Janus.Windows.EditControls.UIButton btnMoveDown;
		internal Janus.Windows.EditControls.UIButton btnMoveUp;
		internal Janus.Windows.EditControls.UIButton btnRemove;
		internal Janus.Windows.EditControls.UIButton btnAdd;
		internal Janus.Windows.GridEX.GridEX grdCompositeColumns;
		internal Janus.Windows.EditControls.UIComboBox cboTable;
		internal Janus.Windows.GridEX.GridEX grdColumnList;
		internal System.Windows.Forms.Label lblSelectTable;
		internal System.Windows.Forms.Label lblCompositeColumns;
		internal System.Windows.Forms.Label lblAvailableColumns;
		internal Janus.Windows.EditControls.UIGroupBox grbTable;
		internal System.Windows.Forms.ImageList ImageList1;


        //#endregion
    }

} //end of root namespace