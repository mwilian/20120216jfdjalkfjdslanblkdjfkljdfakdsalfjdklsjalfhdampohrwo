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
	public partial class ColumnGroupControl : System.Windows.Forms.UserControl
	{

		//UserControl overrides dispose to clean up the component list.
		internal ColumnGroupControl()
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
            this.grbTable = new Janus.Windows.EditControls.UIGroupBox();
            this.lblTables = new System.Windows.Forms.Label();
            this.cboTables = new Janus.Windows.EditControls.UIComboBox();
            this.lblGroupInterval = new System.Windows.Forms.Label();
            this.lblGroupBy = new System.Windows.Forms.Label();
            this.cboColumns = new Janus.Windows.EditControls.UIComboBox();
            this.cboGroupInterval = new Janus.Windows.EditControls.UIComboBox();
            this.grbColumn = new Janus.Windows.EditControls.UIGroupBox();
            this.optDescending = new Janus.Windows.EditControls.UIRadioButton();
            this.optAscending = new Janus.Windows.EditControls.UIRadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.grbTable)).BeginInit();
            this.grbTable.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grbColumn)).BeginInit();
            this.grbColumn.SuspendLayout();
            this.SuspendLayout();
            // 
            // grbTable
            // 
            this.grbTable.BackgroundStyle = Janus.Windows.EditControls.BackgroundStyle.Panel;
            this.grbTable.Controls.Add(this.lblTables);
            this.grbTable.Controls.Add(this.cboTables);
            this.grbTable.Dock = System.Windows.Forms.DockStyle.Top;
            this.grbTable.FrameStyle = Janus.Windows.EditControls.FrameStyle.None;
            this.grbTable.Location = new System.Drawing.Point(0, 0);
            this.grbTable.Name = "grbTable";
            this.grbTable.Size = new System.Drawing.Size(432, 32);
            this.grbTable.TabIndex = 2;
            this.grbTable.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // lblTables
            // 
            this.lblTables.AutoSize = true;
            this.lblTables.BackColor = System.Drawing.Color.Transparent;
            this.lblTables.Location = new System.Drawing.Point(12, 13);
            this.lblTables.Name = "lblTables";
            this.lblTables.Size = new System.Drawing.Size(65, 13);
            this.lblTables.TabIndex = 5;
            this.lblTables.Text = "Select from:";
            // 
            // cboTables
            // 
            this.cboTables.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
            this.cboTables.Location = new System.Drawing.Point(81, 9);
            this.cboTables.Name = "cboTables";
            this.cboTables.Size = new System.Drawing.Size(194, 21);
            this.cboTables.TabIndex = 4;
            this.cboTables.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.cboTables.SelectedItemChanged += new System.EventHandler(this.cboTables_SelectedItemChanged);
            // 
            // lblGroupInterval
            // 
            this.lblGroupInterval.AutoSize = true;
            this.lblGroupInterval.BackColor = System.Drawing.Color.Transparent;
            this.lblGroupInterval.Location = new System.Drawing.Point(12, 71);
            this.lblGroupInterval.Name = "lblGroupInterval";
            this.lblGroupInterval.Size = new System.Drawing.Size(45, 13);
            this.lblGroupInterval.TabIndex = 9;
            this.lblGroupInterval.Text = "Interval";
            this.lblGroupInterval.Visible = false;
            // 
            // lblGroupBy
            // 
            this.lblGroupBy.AutoSize = true;
            this.lblGroupBy.BackColor = System.Drawing.Color.Transparent;
            this.lblGroupBy.Location = new System.Drawing.Point(12, 13);
            this.lblGroupBy.Name = "lblGroupBy";
            this.lblGroupBy.Size = new System.Drawing.Size(51, 13);
            this.lblGroupBy.TabIndex = 8;
            this.lblGroupBy.Text = "Group By";
            // 
            // cboColumns
            // 
            this.cboColumns.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
            this.cboColumns.Location = new System.Drawing.Point(81, 9);
            this.cboColumns.Name = "cboColumns";
            this.cboColumns.Size = new System.Drawing.Size(194, 21);
            this.cboColumns.TabIndex = 7;
            this.cboColumns.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.cboColumns.SelectedItemChanged += new System.EventHandler(this.cboColumns_SelectedItemChanged);
            // 
            // cboGroupInterval
            // 
            this.cboGroupInterval.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
            this.cboGroupInterval.Location = new System.Drawing.Point(81, 67);
            this.cboGroupInterval.Name = "cboGroupInterval";
            this.cboGroupInterval.Size = new System.Drawing.Size(194, 21);
            this.cboGroupInterval.TabIndex = 6;
            this.cboGroupInterval.Visible = false;
            this.cboGroupInterval.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.cboGroupInterval.SelectedItemChanged += new System.EventHandler(this.cboGroupInterval_SelectedItemChanged);
            // 
            // grbColumn
            // 
            this.grbColumn.BackgroundStyle = Janus.Windows.EditControls.BackgroundStyle.Panel;
            this.grbColumn.Controls.Add(this.optDescending);
            this.grbColumn.Controls.Add(this.optAscending);
            this.grbColumn.Controls.Add(this.lblGroupInterval);
            this.grbColumn.Controls.Add(this.lblGroupBy);
            this.grbColumn.Controls.Add(this.cboGroupInterval);
            this.grbColumn.Controls.Add(this.cboColumns);
            this.grbColumn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grbColumn.FrameStyle = Janus.Windows.EditControls.FrameStyle.None;
            this.grbColumn.Location = new System.Drawing.Point(0, 32);
            this.grbColumn.Name = "grbColumn";
            this.grbColumn.Size = new System.Drawing.Size(432, 220);
            this.grbColumn.TabIndex = 10;
            this.grbColumn.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // optDescending
            // 
            this.optDescending.BackColor = System.Drawing.Color.Transparent;
            this.optDescending.Location = new System.Drawing.Point(187, 40);
            this.optDescending.Name = "optDescending";
            this.optDescending.Size = new System.Drawing.Size(88, 16);
            this.optDescending.TabIndex = 11;
            this.optDescending.Text = "Descending";
            this.optDescending.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.optDescending.CheckedChanged += new System.EventHandler(this.optDescending_CheckedChanged);
            // 
            // optAscending
            // 
            this.optAscending.BackColor = System.Drawing.Color.Transparent;
            this.optAscending.Checked = true;
            this.optAscending.Location = new System.Drawing.Point(81, 40);
            this.optAscending.Name = "optAscending";
            this.optAscending.Size = new System.Drawing.Size(88, 16);
            this.optAscending.TabIndex = 10;
            this.optAscending.TabStop = true;
            this.optAscending.Text = "Ascending";
            this.optAscending.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.optAscending.CheckedChanged += new System.EventHandler(this.optAscending_CheckedChanged);
            // 
            // ColumnGroupControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.grbColumn);
            this.Controls.Add(this.grbTable);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "ColumnGroupControl";
            this.Size = new System.Drawing.Size(432, 252);
            ((System.ComponentModel.ISupportInitialize)(this.grbTable)).EndInit();
            this.grbTable.ResumeLayout(false);
            this.grbTable.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grbColumn)).EndInit();
            this.grbColumn.ResumeLayout(false);
            this.grbColumn.PerformLayout();
            this.ResumeLayout(false);

		}
		internal Janus.Windows.EditControls.UIGroupBox grbTable;
		internal System.Windows.Forms.Label lblTables;
		internal Janus.Windows.EditControls.UIComboBox cboTables;
		internal System.Windows.Forms.Label lblGroupInterval;
		internal System.Windows.Forms.Label lblGroupBy;
		internal Janus.Windows.EditControls.UIComboBox cboColumns;
		internal Janus.Windows.EditControls.UIComboBox cboGroupInterval;
		internal Janus.Windows.EditControls.UIGroupBox grbColumn;
		internal Janus.Windows.EditControls.UIRadioButton optDescending;
		internal Janus.Windows.EditControls.UIRadioButton optAscending;


	}

} //end of root namespace