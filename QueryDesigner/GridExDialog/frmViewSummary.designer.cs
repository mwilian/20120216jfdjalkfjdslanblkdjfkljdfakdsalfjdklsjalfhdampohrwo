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
	public partial class frmViewSummary : System.Windows.Forms.Form
	{

		//Form overrides dispose to clean up the component list.
		internal frmViewSummary()
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
		internal Janus.Windows.EditControls.UIGroupBox grbBackground;
		internal Janus.Windows.EditControls.UIButton btnOK;
		internal Janus.Windows.EditControls.UIGroupBox GroupBox1;
		internal Janus.Windows.EditControls.UIButton btnAutoFormatting;
		internal System.Windows.Forms.Label lblFormat;
		internal Janus.Windows.EditControls.UIButton btnFormat;
		internal System.Windows.Forms.Label lblSort;
		internal Janus.Windows.EditControls.UIButton btnSort;
		internal System.Windows.Forms.Label lblGroupBy;
		internal Janus.Windows.EditControls.UIButton btnGroupBy;
		internal System.Windows.Forms.Label lblFields;
		internal Janus.Windows.EditControls.UIButton btnFields;
		internal System.Windows.Forms.Label Label2;
		private void InitializeComponent()
		{
            this.grbBackground = new Janus.Windows.EditControls.UIGroupBox();
            this.btnOK = new Janus.Windows.EditControls.UIButton();
            this.GroupBox1 = new Janus.Windows.EditControls.UIGroupBox();
            this.lblFilterBy = new System.Windows.Forms.Label();
            this.btnFilterBy = new Janus.Windows.EditControls.UIButton();
            this.btnAutoFormatting = new Janus.Windows.EditControls.UIButton();
            this.lblFormat = new System.Windows.Forms.Label();
            this.btnFormat = new Janus.Windows.EditControls.UIButton();
            this.lblSort = new System.Windows.Forms.Label();
            this.btnSort = new Janus.Windows.EditControls.UIButton();
            this.lblGroupBy = new System.Windows.Forms.Label();
            this.btnGroupBy = new Janus.Windows.EditControls.UIButton();
            this.lblFields = new System.Windows.Forms.Label();
            this.btnFields = new Janus.Windows.EditControls.UIButton();
            this.Label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.grbBackground)).BeginInit();
            this.grbBackground.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GroupBox1)).BeginInit();
            this.GroupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // grbBackground
            // 
            this.grbBackground.BackgroundStyle = Janus.Windows.EditControls.BackgroundStyle.Panel;
            this.grbBackground.Controls.Add(this.btnOK);
            this.grbBackground.Controls.Add(this.GroupBox1);
            this.grbBackground.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grbBackground.FrameStyle = Janus.Windows.EditControls.FrameStyle.None;
            this.grbBackground.Location = new System.Drawing.Point(0, 0);
            this.grbBackground.Name = "grbBackground";
            this.grbBackground.Size = new System.Drawing.Size(499, 264);
            this.grbBackground.TabIndex = 0;
            this.grbBackground.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOK.Location = new System.Drawing.Point(408, 231);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(80, 24);
            this.btnOK.TabIndex = 10;
            this.btnOK.Text = "OK";
            this.btnOK.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // GroupBox1
            // 
            this.GroupBox1.BackColor = System.Drawing.Color.Transparent;
            this.GroupBox1.Controls.Add(this.lblFilterBy);
            this.GroupBox1.Controls.Add(this.btnFilterBy);
            this.GroupBox1.Controls.Add(this.btnAutoFormatting);
            this.GroupBox1.Controls.Add(this.lblFormat);
            this.GroupBox1.Controls.Add(this.btnFormat);
            this.GroupBox1.Controls.Add(this.lblSort);
            this.GroupBox1.Controls.Add(this.btnSort);
            this.GroupBox1.Controls.Add(this.lblGroupBy);
            this.GroupBox1.Controls.Add(this.btnGroupBy);
            this.GroupBox1.Controls.Add(this.lblFields);
            this.GroupBox1.Controls.Add(this.btnFields);
            this.GroupBox1.Controls.Add(this.Label2);
            this.GroupBox1.Location = new System.Drawing.Point(8, 8);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(480, 217);
            this.GroupBox1.TabIndex = 12;
            this.GroupBox1.Text = "Description";
            this.GroupBox1.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // lblFilterBy
            // 
            this.lblFilterBy.Location = new System.Drawing.Point(143, 84);
            this.lblFilterBy.Name = "lblFilterBy";
            this.lblFilterBy.Size = new System.Drawing.Size(327, 28);
            this.lblFilterBy.TabIndex = 11;
            this.lblFilterBy.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnFilterBy
            // 
            this.btnFilterBy.Location = new System.Drawing.Point(8, 86);
            this.btnFilterBy.Name = "btnFilterBy";
            this.btnFilterBy.Size = new System.Drawing.Size(128, 24);
            this.btnFilterBy.TabIndex = 10;
            this.btnFilterBy.Text = "Filter By...";
            this.btnFilterBy.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnFilterBy.Click += new System.EventHandler(this.btnFilterBy_Click);
            // 
            // btnAutoFormatting
            // 
            this.btnAutoFormatting.Location = new System.Drawing.Point(8, 182);
            this.btnAutoFormatting.Name = "btnAutoFormatting";
            this.btnAutoFormatting.Size = new System.Drawing.Size(128, 24);
            this.btnAutoFormatting.TabIndex = 8;
            this.btnAutoFormatting.Text = "Automatic Formatting...";
            this.btnAutoFormatting.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnAutoFormatting.Click += new System.EventHandler(this.btnAutoFormatting_Click);
            // 
            // lblFormat
            // 
            this.lblFormat.Location = new System.Drawing.Point(143, 148);
            this.lblFormat.Name = "lblFormat";
            this.lblFormat.Size = new System.Drawing.Size(327, 28);
            this.lblFormat.TabIndex = 7;
            this.lblFormat.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnFormat
            // 
            this.btnFormat.Location = new System.Drawing.Point(8, 150);
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Size = new System.Drawing.Size(128, 24);
            this.btnFormat.TabIndex = 6;
            this.btnFormat.Text = "Format...";
            this.btnFormat.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnFormat.Click += new System.EventHandler(this.btnFormat_Click);
            // 
            // lblSort
            // 
            this.lblSort.Location = new System.Drawing.Point(143, 116);
            this.lblSort.Name = "lblSort";
            this.lblSort.Size = new System.Drawing.Size(327, 28);
            this.lblSort.TabIndex = 5;
            this.lblSort.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnSort
            // 
            this.btnSort.Location = new System.Drawing.Point(8, 118);
            this.btnSort.Name = "btnSort";
            this.btnSort.Size = new System.Drawing.Size(128, 24);
            this.btnSort.TabIndex = 4;
            this.btnSort.Text = "Sort...";
            this.btnSort.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnSort.Click += new System.EventHandler(this.btnSort_Click);
            // 
            // lblGroupBy
            // 
            this.lblGroupBy.Location = new System.Drawing.Point(143, 52);
            this.lblGroupBy.Name = "lblGroupBy";
            this.lblGroupBy.Size = new System.Drawing.Size(327, 28);
            this.lblGroupBy.TabIndex = 3;
            this.lblGroupBy.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnGroupBy
            // 
            this.btnGroupBy.Location = new System.Drawing.Point(8, 54);
            this.btnGroupBy.Name = "btnGroupBy";
            this.btnGroupBy.Size = new System.Drawing.Size(128, 24);
            this.btnGroupBy.TabIndex = 2;
            this.btnGroupBy.Text = "Group By...";
            this.btnGroupBy.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnGroupBy.Click += new System.EventHandler(this.btnGroupBy_Click);
            // 
            // lblFields
            // 
            this.lblFields.Location = new System.Drawing.Point(143, 18);
            this.lblFields.Name = "lblFields";
            this.lblFields.Size = new System.Drawing.Size(327, 28);
            this.lblFields.TabIndex = 1;
            this.lblFields.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnFields
            // 
            this.btnFields.Location = new System.Drawing.Point(8, 20);
            this.btnFields.Name = "btnFields";
            this.btnFields.Size = new System.Drawing.Size(128, 24);
            this.btnFields.TabIndex = 0;
            this.btnFields.Text = "Fields...";
            this.btnFields.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnFields.Click += new System.EventHandler(this.btnFields_Click);
            // 
            // Label2
            // 
            this.Label2.Location = new System.Drawing.Point(142, 180);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(327, 28);
            this.Label2.TabIndex = 9;
            this.Label2.Text = "Condition font and color formatting.";
            this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // frmViewSummary
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
            this.ClientSize = new System.Drawing.Size(499, 264);
            this.Controls.Add(this.grbBackground);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmViewSummary";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "View Summary";
            ((System.ComponentModel.ISupportInitialize)(this.grbBackground)).EndInit();
            this.grbBackground.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.GroupBox1)).EndInit();
            this.GroupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		internal System.Windows.Forms.Label lblFilterBy;
		internal Janus.Windows.EditControls.UIButton btnFilterBy;

	}

} //end of root namespace