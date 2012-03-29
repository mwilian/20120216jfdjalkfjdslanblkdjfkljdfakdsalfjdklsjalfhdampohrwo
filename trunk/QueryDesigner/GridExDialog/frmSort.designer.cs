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
	public partial class frmSort : System.Windows.Forms.Form
	{

		//Form overrides dispose to clean up the component list.
		internal frmSort()
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
        internal GroupBox grbBackground;
        internal Button btnClear;
        internal ComboBox cboColumn3;
        internal RadioButton optDescending3;
        internal RadioButton optAscending3;
        internal Button btnOK;
        internal ComboBox cboColumn1;
        internal RadioButton optDescending1;
        internal RadioButton optAscending1;
        internal ComboBox cboColumn0;
        internal RadioButton optDescending0;
        internal RadioButton optAscending0;
        internal Button btnCancel;
        internal ComboBox cboColumn2;
        internal RadioButton optDescending2;
        internal RadioButton optAscending2;
        internal GroupBox grbSort4;
        internal GroupBox grbSort2;
        internal GroupBox grbSort1;
        internal GroupBox grbSort3;
		private void InitializeComponent()
		{
            this.grbBackground = new System.Windows.Forms.GroupBox();
            this.btnClear = new System.Windows.Forms.Button();
            this.grbSort4 = new System.Windows.Forms.GroupBox();
            this.cboColumn3 = new System.Windows.Forms.ComboBox();
            this.optDescending3 = new System.Windows.Forms.RadioButton();
            this.optAscending3 = new System.Windows.Forms.RadioButton();
            this.btnOK = new System.Windows.Forms.Button();
            this.grbSort2 = new System.Windows.Forms.GroupBox();
            this.cboColumn1 = new System.Windows.Forms.ComboBox();
            this.optDescending1 = new System.Windows.Forms.RadioButton();
            this.optAscending1 = new System.Windows.Forms.RadioButton();
            this.grbSort1 = new System.Windows.Forms.GroupBox();
            this.cboColumn0 = new System.Windows.Forms.ComboBox();
            this.optDescending0 = new System.Windows.Forms.RadioButton();
            this.optAscending0 = new System.Windows.Forms.RadioButton();
            this.btnCancel = new System.Windows.Forms.Button();
            this.grbSort3 = new System.Windows.Forms.GroupBox();
            this.cboColumn2 = new System.Windows.Forms.ComboBox();
            this.optDescending2 = new System.Windows.Forms.RadioButton();
            this.optAscending2 = new System.Windows.Forms.RadioButton();
            this.grbBackground.SuspendLayout();
            this.grbSort4.SuspendLayout();
            this.grbSort2.SuspendLayout();
            this.grbSort1.SuspendLayout();
            this.grbSort3.SuspendLayout();
            this.SuspendLayout();
            // 
            // grbBackground
            // 
            this.grbBackground.Controls.Add(this.btnClear);
            this.grbBackground.Controls.Add(this.grbSort4);
            this.grbBackground.Controls.Add(this.btnOK);
            this.grbBackground.Controls.Add(this.grbSort2);
            this.grbBackground.Controls.Add(this.grbSort1);
            this.grbBackground.Controls.Add(this.btnCancel);
            this.grbBackground.Controls.Add(this.grbSort3);
            this.grbBackground.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grbBackground.Location = new System.Drawing.Point(0, 0);
            this.grbBackground.Name = "grbBackground";
            this.grbBackground.Size = new System.Drawing.Size(402, 304);
            this.grbBackground.TabIndex = 0;
            this.grbBackground.TabStop = false;
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(312, 72);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(80, 24);
            this.btnClear.TabIndex = 13;
            this.btnClear.Text = "Clear All";
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // grbSort4
            // 
            this.grbSort4.BackColor = System.Drawing.Color.Transparent;
            this.grbSort4.Controls.Add(this.cboColumn3);
            this.grbSort4.Controls.Add(this.optDescending3);
            this.grbSort4.Controls.Add(this.optAscending3);
            this.grbSort4.Location = new System.Drawing.Point(8, 230);
            this.grbSort4.Name = "grbSort4";
            this.grbSort4.Size = new System.Drawing.Size(296, 64);
            this.grbSort4.TabIndex = 14;
            this.grbSort4.TabStop = false;
            this.grbSort4.Text = "Then By";
            // 
            // cboColumn3
            // 
            this.cboColumn3.Location = new System.Drawing.Point(8, 16);
            this.cboColumn3.Name = "cboColumn3";
            this.cboColumn3.Size = new System.Drawing.Size(176, 21);
            this.cboColumn3.TabIndex = 4;
            this.cboColumn3.SelectedValueChanged += new System.EventHandler(this.cboColumn3_SelectedItemChanged);
            // 
            // optDescending3
            // 
            this.optDescending3.Enabled = false;
            this.optDescending3.Location = new System.Drawing.Point(196, 36);
            this.optDescending3.Name = "optDescending3";
            this.optDescending3.Size = new System.Drawing.Size(88, 16);
            this.optDescending3.TabIndex = 2;
            this.optDescending3.Text = "Descending";
            // 
            // optAscending3
            // 
            this.optAscending3.Checked = true;
            this.optAscending3.Enabled = false;
            this.optAscending3.Location = new System.Drawing.Point(196, 16);
            this.optAscending3.Name = "optAscending3";
            this.optAscending3.Size = new System.Drawing.Size(88, 16);
            this.optAscending3.TabIndex = 1;
            this.optAscending3.TabStop = true;
            this.optAscending3.Text = "Ascending";
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(312, 8);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(80, 24);
            this.btnOK.TabIndex = 11;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // grbSort2
            // 
            this.grbSort2.BackColor = System.Drawing.Color.Transparent;
            this.grbSort2.Controls.Add(this.cboColumn1);
            this.grbSort2.Controls.Add(this.optDescending1);
            this.grbSort2.Controls.Add(this.optAscending1);
            this.grbSort2.Location = new System.Drawing.Point(8, 82);
            this.grbSort2.Name = "grbSort2";
            this.grbSort2.Size = new System.Drawing.Size(296, 64);
            this.grbSort2.TabIndex = 15;
            this.grbSort2.TabStop = false;
            this.grbSort2.Text = "Then By";
            // 
            // cboColumn1
            // 
            this.cboColumn1.Location = new System.Drawing.Point(8, 16);
            this.cboColumn1.Name = "cboColumn1";
            this.cboColumn1.Size = new System.Drawing.Size(176, 21);
            this.cboColumn1.TabIndex = 4;
            this.cboColumn1.SelectedValueChanged += new System.EventHandler(this.cboColumn1_SelectedItemChanged);
            // 
            // optDescending1
            // 
            this.optDescending1.Enabled = false;
            this.optDescending1.Location = new System.Drawing.Point(196, 36);
            this.optDescending1.Name = "optDescending1";
            this.optDescending1.Size = new System.Drawing.Size(88, 16);
            this.optDescending1.TabIndex = 2;
            this.optDescending1.Text = "Descending";
            // 
            // optAscending1
            // 
            this.optAscending1.Checked = true;
            this.optAscending1.Enabled = false;
            this.optAscending1.Location = new System.Drawing.Point(196, 16);
            this.optAscending1.Name = "optAscending1";
            this.optAscending1.Size = new System.Drawing.Size(88, 16);
            this.optAscending1.TabIndex = 1;
            this.optAscending1.TabStop = true;
            this.optAscending1.Text = "Ascending";
            // 
            // grbSort1
            // 
            this.grbSort1.BackColor = System.Drawing.Color.Transparent;
            this.grbSort1.Controls.Add(this.cboColumn0);
            this.grbSort1.Controls.Add(this.optDescending0);
            this.grbSort1.Controls.Add(this.optAscending0);
            this.grbSort1.Location = new System.Drawing.Point(8, 8);
            this.grbSort1.Name = "grbSort1";
            this.grbSort1.Size = new System.Drawing.Size(296, 64);
            this.grbSort1.TabIndex = 16;
            this.grbSort1.TabStop = false;
            this.grbSort1.Text = "Sort Items By";
            // 
            // cboColumn0
            // 
            this.cboColumn0.Location = new System.Drawing.Point(8, 16);
            this.cboColumn0.Name = "cboColumn0";
            this.cboColumn0.Size = new System.Drawing.Size(176, 21);
            this.cboColumn0.TabIndex = 3;
            this.cboColumn0.SelectedValueChanged += new System.EventHandler(this.cboColumn0_SelectedItemChanged);
            // 
            // optDescending0
            // 
            this.optDescending0.Enabled = false;
            this.optDescending0.Location = new System.Drawing.Point(196, 36);
            this.optDescending0.Name = "optDescending0";
            this.optDescending0.Size = new System.Drawing.Size(88, 16);
            this.optDescending0.TabIndex = 2;
            this.optDescending0.Text = "Descending";
            // 
            // optAscending0
            // 
            this.optAscending0.Checked = true;
            this.optAscending0.Enabled = false;
            this.optAscending0.Location = new System.Drawing.Point(196, 16);
            this.optAscending0.Name = "optAscending0";
            this.optAscending0.Size = new System.Drawing.Size(88, 16);
            this.optAscending0.TabIndex = 1;
            this.optAscending0.TabStop = true;
            this.optAscending0.Text = "Ascending";
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(312, 40);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(80, 24);
            this.btnCancel.TabIndex = 12;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // grbSort3
            // 
            this.grbSort3.BackColor = System.Drawing.Color.Transparent;
            this.grbSort3.Controls.Add(this.cboColumn2);
            this.grbSort3.Controls.Add(this.optDescending2);
            this.grbSort3.Controls.Add(this.optAscending2);
            this.grbSort3.Location = new System.Drawing.Point(8, 156);
            this.grbSort3.Name = "grbSort3";
            this.grbSort3.Size = new System.Drawing.Size(296, 64);
            this.grbSort3.TabIndex = 17;
            this.grbSort3.TabStop = false;
            this.grbSort3.Text = "Then By";
            // 
            // cboColumn2
            // 
            this.cboColumn2.Location = new System.Drawing.Point(8, 16);
            this.cboColumn2.Name = "cboColumn2";
            this.cboColumn2.Size = new System.Drawing.Size(176, 21);
            this.cboColumn2.TabIndex = 4;
            this.cboColumn2.SelectedValueChanged += new System.EventHandler(this.cboColumn2_SelectedItemChanged);
            // 
            // optDescending2
            // 
            this.optDescending2.Enabled = false;
            this.optDescending2.Location = new System.Drawing.Point(196, 36);
            this.optDescending2.Name = "optDescending2";
            this.optDescending2.Size = new System.Drawing.Size(88, 16);
            this.optDescending2.TabIndex = 2;
            this.optDescending2.Text = "Descending";
            // 
            // optAscending2
            // 
            this.optAscending2.Checked = true;
            this.optAscending2.Enabled = false;
            this.optAscending2.Location = new System.Drawing.Point(196, 16);
            this.optAscending2.Name = "optAscending2";
            this.optAscending2.Size = new System.Drawing.Size(88, 16);
            this.optAscending2.TabIndex = 1;
            this.optAscending2.TabStop = true;
            this.optAscending2.Text = "Ascending";
            // 
            // frmSort
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
            this.ClientSize = new System.Drawing.Size(402, 304);
            this.Controls.Add(this.grbBackground);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmSort";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Sort";
            this.grbBackground.ResumeLayout(false);
            this.grbSort4.ResumeLayout(false);
            this.grbSort2.ResumeLayout(false);
            this.grbSort1.ResumeLayout(false);
            this.grbSort3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

	}

} //end of root namespace