using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;

using Janus.Windows.GridEX;
using Janus.Windows.EditControls;

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
		internal Janus.Windows.EditControls.UIGroupBox grbBackground;
		internal Janus.Windows.EditControls.UIButton btnClear;
		internal Janus.Windows.EditControls.UIComboBox cboColumn3;
		internal Janus.Windows.EditControls.UIRadioButton optDescending3;
		internal Janus.Windows.EditControls.UIRadioButton optAscending3;
		internal Janus.Windows.EditControls.UIButton btnOK;
		internal Janus.Windows.EditControls.UIComboBox cboColumn1;
		internal Janus.Windows.EditControls.UIRadioButton optDescending1;
		internal Janus.Windows.EditControls.UIRadioButton optAscending1;
		internal Janus.Windows.EditControls.UIComboBox cboColumn0;
		internal Janus.Windows.EditControls.UIRadioButton optDescending0;
		internal Janus.Windows.EditControls.UIRadioButton optAscending0;
		internal Janus.Windows.EditControls.UIButton btnCancel;
		internal Janus.Windows.EditControls.UIComboBox cboColumn2;
		internal Janus.Windows.EditControls.UIRadioButton optDescending2;
		internal Janus.Windows.EditControls.UIRadioButton optAscending2;
		internal Janus.Windows.EditControls.UIGroupBox grbSort4;
		internal Janus.Windows.EditControls.UIGroupBox grbSort2;
		internal Janus.Windows.EditControls.UIGroupBox grbSort1;
		internal Janus.Windows.EditControls.UIGroupBox grbSort3;
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.grbBackground = new Janus.Windows.EditControls.UIGroupBox();
            this.btnClear = new Janus.Windows.EditControls.UIButton();
            this.grbSort4 = new Janus.Windows.EditControls.UIGroupBox();
            this.cboColumn3 = new Janus.Windows.EditControls.UIComboBox();
            this.optDescending3 = new Janus.Windows.EditControls.UIRadioButton();
            this.optAscending3 = new Janus.Windows.EditControls.UIRadioButton();
            this.btnOK = new Janus.Windows.EditControls.UIButton();
            this.grbSort2 = new Janus.Windows.EditControls.UIGroupBox();
            this.cboColumn1 = new Janus.Windows.EditControls.UIComboBox();
            this.optDescending1 = new Janus.Windows.EditControls.UIRadioButton();
            this.optAscending1 = new Janus.Windows.EditControls.UIRadioButton();
            this.grbSort1 = new Janus.Windows.EditControls.UIGroupBox();
            this.cboColumn0 = new Janus.Windows.EditControls.UIComboBox();
            this.optDescending0 = new Janus.Windows.EditControls.UIRadioButton();
            this.optAscending0 = new Janus.Windows.EditControls.UIRadioButton();
            this.btnCancel = new Janus.Windows.EditControls.UIButton();
            this.grbSort3 = new Janus.Windows.EditControls.UIGroupBox();
            this.cboColumn2 = new Janus.Windows.EditControls.UIComboBox();
            this.optDescending2 = new Janus.Windows.EditControls.UIRadioButton();
            this.optAscending2 = new Janus.Windows.EditControls.UIRadioButton();
            this.officeFormAdorner1 = new Janus.Windows.Ribbon.OfficeFormAdorner(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.grbBackground)).BeginInit();
            this.grbBackground.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grbSort4)).BeginInit();
            this.grbSort4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grbSort2)).BeginInit();
            this.grbSort2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grbSort1)).BeginInit();
            this.grbSort1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grbSort3)).BeginInit();
            this.grbSort3.SuspendLayout();
            this.SuspendLayout();
            // 
            // grbBackground
            // 
            this.grbBackground.BackgroundStyle = Janus.Windows.EditControls.BackgroundStyle.Panel;
            this.grbBackground.Controls.Add(this.btnClear);
            this.grbBackground.Controls.Add(this.grbSort4);
            this.grbBackground.Controls.Add(this.btnOK);
            this.grbBackground.Controls.Add(this.grbSort2);
            this.grbBackground.Controls.Add(this.grbSort1);
            this.grbBackground.Controls.Add(this.btnCancel);
            this.grbBackground.Controls.Add(this.grbSort3);
            this.grbBackground.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grbBackground.FrameStyle = Janus.Windows.EditControls.FrameStyle.None;
            this.grbBackground.Location = new System.Drawing.Point(0, 0);
            this.grbBackground.Name = "grbBackground";
            this.grbBackground.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.grbBackground.Office2007CustomColor = System.Drawing.Color.Empty;
            this.grbBackground.Size = new System.Drawing.Size(400, 302);
            this.grbBackground.TabIndex = 0;
            this.grbBackground.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(312, 72);
            this.btnClear.Name = "btnClear";
            this.btnClear.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.btnClear.Office2007CustomColor = System.Drawing.Color.Empty;
            this.btnClear.Size = new System.Drawing.Size(80, 24);
            this.btnClear.TabIndex = 13;
            this.btnClear.Text = "Clear All";
            this.btnClear.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
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
            this.grbSort4.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.grbSort4.Office2007CustomColor = System.Drawing.Color.Empty;
            this.grbSort4.Size = new System.Drawing.Size(296, 64);
            this.grbSort4.TabIndex = 14;
            this.grbSort4.Text = "Then By";
            this.grbSort4.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // cboColumn3
            // 
            this.cboColumn3.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
            this.cboColumn3.Location = new System.Drawing.Point(8, 16);
            this.cboColumn3.Name = "cboColumn3";
            this.cboColumn3.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.cboColumn3.Office2007CustomColor = System.Drawing.Color.Empty;
            this.cboColumn3.Size = new System.Drawing.Size(176, 21);
            this.cboColumn3.TabIndex = 4;
            this.cboColumn3.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.cboColumn3.SelectedItemChanged += new System.EventHandler(this.cboColumn3_SelectedItemChanged);
            // 
            // optDescending3
            // 
            this.optDescending3.Enabled = false;
            this.optDescending3.Location = new System.Drawing.Point(196, 36);
            this.optDescending3.Name = "optDescending3";
            this.optDescending3.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.optDescending3.Office2007CustomColor = System.Drawing.Color.Empty;
            this.optDescending3.Size = new System.Drawing.Size(88, 16);
            this.optDescending3.TabIndex = 2;
            this.optDescending3.Text = "Descending";
            this.optDescending3.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // optAscending3
            // 
            this.optAscending3.Checked = true;
            this.optAscending3.Enabled = false;
            this.optAscending3.Location = new System.Drawing.Point(196, 16);
            this.optAscending3.Name = "optAscending3";
            this.optAscending3.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.optAscending3.Office2007CustomColor = System.Drawing.Color.Empty;
            this.optAscending3.Size = new System.Drawing.Size(88, 16);
            this.optAscending3.TabIndex = 1;
            this.optAscending3.TabStop = true;
            this.optAscending3.Text = "Ascending";
            this.optAscending3.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(312, 8);
            this.btnOK.Name = "btnOK";
            this.btnOK.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.btnOK.Office2007CustomColor = System.Drawing.Color.Empty;
            this.btnOK.Size = new System.Drawing.Size(80, 24);
            this.btnOK.TabIndex = 11;
            this.btnOK.Text = "OK";
            this.btnOK.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
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
            this.grbSort2.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.grbSort2.Office2007CustomColor = System.Drawing.Color.Empty;
            this.grbSort2.Size = new System.Drawing.Size(296, 64);
            this.grbSort2.TabIndex = 15;
            this.grbSort2.Text = "Then By";
            this.grbSort2.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // cboColumn1
            // 
            this.cboColumn1.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
            this.cboColumn1.Location = new System.Drawing.Point(8, 16);
            this.cboColumn1.Name = "cboColumn1";
            this.cboColumn1.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.cboColumn1.Office2007CustomColor = System.Drawing.Color.Empty;
            this.cboColumn1.Size = new System.Drawing.Size(176, 21);
            this.cboColumn1.TabIndex = 4;
            this.cboColumn1.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.cboColumn1.SelectedItemChanged += new System.EventHandler(this.cboColumn1_SelectedItemChanged);
            // 
            // optDescending1
            // 
            this.optDescending1.Enabled = false;
            this.optDescending1.Location = new System.Drawing.Point(196, 36);
            this.optDescending1.Name = "optDescending1";
            this.optDescending1.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.optDescending1.Office2007CustomColor = System.Drawing.Color.Empty;
            this.optDescending1.Size = new System.Drawing.Size(88, 16);
            this.optDescending1.TabIndex = 2;
            this.optDescending1.Text = "Descending";
            this.optDescending1.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // optAscending1
            // 
            this.optAscending1.Checked = true;
            this.optAscending1.Enabled = false;
            this.optAscending1.Location = new System.Drawing.Point(196, 16);
            this.optAscending1.Name = "optAscending1";
            this.optAscending1.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.optAscending1.Office2007CustomColor = System.Drawing.Color.Empty;
            this.optAscending1.Size = new System.Drawing.Size(88, 16);
            this.optAscending1.TabIndex = 1;
            this.optAscending1.TabStop = true;
            this.optAscending1.Text = "Ascending";
            this.optAscending1.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // grbSort1
            // 
            this.grbSort1.BackColor = System.Drawing.Color.Transparent;
            this.grbSort1.Controls.Add(this.cboColumn0);
            this.grbSort1.Controls.Add(this.optDescending0);
            this.grbSort1.Controls.Add(this.optAscending0);
            this.grbSort1.Location = new System.Drawing.Point(8, 8);
            this.grbSort1.Name = "grbSort1";
            this.grbSort1.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.grbSort1.Office2007CustomColor = System.Drawing.Color.Empty;
            this.grbSort1.Size = new System.Drawing.Size(296, 64);
            this.grbSort1.TabIndex = 16;
            this.grbSort1.Text = "Sort Items By";
            this.grbSort1.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // cboColumn0
            // 
            this.cboColumn0.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
            this.cboColumn0.Location = new System.Drawing.Point(8, 16);
            this.cboColumn0.Name = "cboColumn0";
            this.cboColumn0.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.cboColumn0.Office2007CustomColor = System.Drawing.Color.Empty;
            this.cboColumn0.Size = new System.Drawing.Size(176, 21);
            this.cboColumn0.TabIndex = 3;
            this.cboColumn0.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.cboColumn0.SelectedItemChanged += new System.EventHandler(this.cboColumn0_SelectedItemChanged);
            // 
            // optDescending0
            // 
            this.optDescending0.Enabled = false;
            this.optDescending0.Location = new System.Drawing.Point(196, 36);
            this.optDescending0.Name = "optDescending0";
            this.optDescending0.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.optDescending0.Office2007CustomColor = System.Drawing.Color.Empty;
            this.optDescending0.Size = new System.Drawing.Size(88, 16);
            this.optDescending0.TabIndex = 2;
            this.optDescending0.Text = "Descending";
            this.optDescending0.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // optAscending0
            // 
            this.optAscending0.Checked = true;
            this.optAscending0.Enabled = false;
            this.optAscending0.Location = new System.Drawing.Point(196, 16);
            this.optAscending0.Name = "optAscending0";
            this.optAscending0.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.optAscending0.Office2007CustomColor = System.Drawing.Color.Empty;
            this.optAscending0.Size = new System.Drawing.Size(88, 16);
            this.optAscending0.TabIndex = 1;
            this.optAscending0.TabStop = true;
            this.optAscending0.Text = "Ascending";
            this.optAscending0.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(312, 40);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.btnCancel.Office2007CustomColor = System.Drawing.Color.Empty;
            this.btnCancel.Size = new System.Drawing.Size(80, 24);
            this.btnCancel.TabIndex = 12;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
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
            this.grbSort3.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.grbSort3.Office2007CustomColor = System.Drawing.Color.Empty;
            this.grbSort3.Size = new System.Drawing.Size(296, 64);
            this.grbSort3.TabIndex = 17;
            this.grbSort3.Text = "Then By";
            this.grbSort3.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // cboColumn2
            // 
            this.cboColumn2.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
            this.cboColumn2.Location = new System.Drawing.Point(8, 16);
            this.cboColumn2.Name = "cboColumn2";
            this.cboColumn2.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.cboColumn2.Office2007CustomColor = System.Drawing.Color.Empty;
            this.cboColumn2.Size = new System.Drawing.Size(176, 21);
            this.cboColumn2.TabIndex = 4;
            this.cboColumn2.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.cboColumn2.SelectedItemChanged += new System.EventHandler(this.cboColumn2_SelectedItemChanged);
            // 
            // optDescending2
            // 
            this.optDescending2.Enabled = false;
            this.optDescending2.Location = new System.Drawing.Point(196, 36);
            this.optDescending2.Name = "optDescending2";
            this.optDescending2.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.optDescending2.Office2007CustomColor = System.Drawing.Color.Empty;
            this.optDescending2.Size = new System.Drawing.Size(88, 16);
            this.optDescending2.TabIndex = 2;
            this.optDescending2.Text = "Descending";
            this.optDescending2.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // optAscending2
            // 
            this.optAscending2.Checked = true;
            this.optAscending2.Enabled = false;
            this.optAscending2.Location = new System.Drawing.Point(196, 16);
            this.optAscending2.Name = "optAscending2";
            this.optAscending2.Office2007ColorScheme = Janus.Windows.UI.Office2007ColorScheme.Default;
            this.optAscending2.Office2007CustomColor = System.Drawing.Color.Empty;
            this.optAscending2.Size = new System.Drawing.Size(88, 16);
            this.optAscending2.TabIndex = 1;
            this.optAscending2.TabStop = true;
            this.optAscending2.Text = "Ascending";
            this.optAscending2.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // officeFormAdorner1
            // 
            this.officeFormAdorner1.DocumentName = "Sort";
            this.officeFormAdorner1.Form = this;
            this.officeFormAdorner1.Office2007CustomColor = System.Drawing.Color.Empty;
            this.officeFormAdorner1.TitleBarFont = null;
            // 
            // frmSort
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
            this.ClientSize = new System.Drawing.Size(400, 302);
            this.Controls.Add(this.grbBackground);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmSort";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Sort";
            ((System.ComponentModel.ISupportInitialize)(this.grbBackground)).EndInit();
            this.grbBackground.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.grbSort4)).EndInit();
            this.grbSort4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.grbSort2)).EndInit();
            this.grbSort2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.grbSort1)).EndInit();
            this.grbSort1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.grbSort3)).EndInit();
            this.grbSort3.ResumeLayout(false);
            this.ResumeLayout(false);

		}

        private Janus.Windows.Ribbon.OfficeFormAdorner officeFormAdorner1;

	}

} //end of root namespace