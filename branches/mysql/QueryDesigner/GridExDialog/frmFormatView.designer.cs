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
	public partial class frmFormatView : System.Windows.Forms.Form
	{

		//Form overrides dispose to clean up the component list.
		internal frmFormatView()
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
		internal System.Windows.Forms.FontDialog FontDialog1;
		internal GroupBox UiGroupBox1;
		internal Button btnCancel;
		internal Button btnOK;
		internal GroupBox GroupBox3;
		internal CheckBox chkShadeGroupHeaders;
		internal ComboBox cboGridlineStyle;
		internal System.Windows.Forms.Label Label2;
		internal GroupBox GroupBox2;
		internal CheckBox chkAllowAddNew;
		internal CheckBox chkAllowEdit;
		internal System.Windows.Forms.Label lblRowsFont;
		internal Button btnRowsFont;
		internal GroupBox GroupBox1;
		internal CheckBox chkAutoSize;
		internal System.Windows.Forms.Label lblHeaderFont;
		internal Button btnHeaderFont;
		private void InitializeComponent()
		{
            this.FontDialog1 = new System.Windows.Forms.FontDialog();
            this.UiGroupBox1 = new GroupBox();
            this.btnCancel = new Button();
            this.btnOK = new Button();
            this.GroupBox3 = new GroupBox();
            this.chkShadeGroupHeaders = new CheckBox();
            this.cboGridlineStyle = new ComboBox();
            this.Label2 = new System.Windows.Forms.Label();
            this.GroupBox2 = new GroupBox();
            this.chkAllowAddNew = new CheckBox();
            this.chkAllowEdit = new CheckBox();
            this.lblRowsFont = new System.Windows.Forms.Label();
            this.btnRowsFont = new Button();
            this.GroupBox1 = new GroupBox();
            this.chkAutoSize = new CheckBox();
            this.lblHeaderFont = new System.Windows.Forms.Label();
            this.btnHeaderFont = new Button();
            //((System.ComponentModel.ISupportInitialize)(this.UiGroupBox1)).BeginInit();
            //this.UiGroupBox1.SuspendLayout();
            //((System.ComponentModel.ISupportInitialize)(this.GroupBox3)).BeginInit();
            //this.GroupBox3.SuspendLayout();
            //((System.ComponentModel.ISupportInitialize)(this.GroupBox2)).BeginInit();
            //this.GroupBox2.SuspendLayout();
            //((System.ComponentModel.ISupportInitialize)(this.GroupBox1)).BeginInit();
            //this.GroupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // FontDialog1
            // 
            this.FontDialog1.AllowVerticalFonts = false;
            this.FontDialog1.FontMustExist = true;
            this.FontDialog1.ShowColor = true;
            // 
            // UiGroupBox1
            // 
            //this.UiGroupBox1.BackgroundStyle = Janus.Windows.EditControls.BackgroundStyle.Panel;
            this.UiGroupBox1.Controls.Add(this.btnCancel);
            this.UiGroupBox1.Controls.Add(this.btnOK);
            this.UiGroupBox1.Controls.Add(this.GroupBox3);
            this.UiGroupBox1.Controls.Add(this.GroupBox2);
            this.UiGroupBox1.Controls.Add(this.GroupBox1);
            this.UiGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.UiGroupBox1.FrameStyle = Janus.Windows.EditControls.FrameStyle.None;
            this.UiGroupBox1.Location = new System.Drawing.Point(0, 0);
            this.UiGroupBox1.Name = "UiGroupBox1";
            this.UiGroupBox1.Size = new System.Drawing.Size(538, 240);
            this.UiGroupBox1.TabIndex = 0;
            //this.UiGroupBox1.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(448, 40);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(80, 24);
            this.btnCancel.TabIndex = 11;
            this.btnCancel.Text = "Cancel";
            //this.btnCancel.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOK.Location = new System.Drawing.Point(448, 8);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(80, 24);
            this.btnOK.TabIndex = 10;
            this.btnOK.Text = "OK";
            //this.btnOK.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // GroupBox3
            // 
            this.GroupBox3.BackColor = System.Drawing.Color.Transparent;
            this.GroupBox3.Controls.Add(this.chkShadeGroupHeaders);
            this.GroupBox3.Controls.Add(this.cboGridlineStyle);
            this.GroupBox3.Controls.Add(this.Label2);
            this.GroupBox3.Location = new System.Drawing.Point(8, 164);
            this.GroupBox3.Name = "GroupBox3";
            this.GroupBox3.Size = new System.Drawing.Size(424, 68);
            this.GroupBox3.TabIndex = 12;
            this.GroupBox3.Text = "Grid lines";
            //this.GroupBox3.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // chkShadeGroupHeaders
            // 
            this.chkShadeGroupHeaders.Location = new System.Drawing.Point(276, 23);
            this.chkShadeGroupHeaders.Name = "chkShadeGroupHeaders";
            this.chkShadeGroupHeaders.Size = new System.Drawing.Size(132, 16);
            this.chkShadeGroupHeaders.TabIndex = 3;
            this.chkShadeGroupHeaders.Text = "Shade group headings";
            // 
            // cboGridlineStyle
            // 
            this.cboGridlineStyle.Location = new System.Drawing.Point(100, 21);
            this.cboGridlineStyle.Name = "cboGridlineStyle";
            this.cboGridlineStyle.Size = new System.Drawing.Size(168, 21);
            this.cboGridlineStyle.TabIndex = 1;
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(12, 24);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(75, 13);
            this.Label2.TabIndex = 0;
            this.Label2.Text = "Grid line style:";
            // 
            // GroupBox2
            // 
            this.GroupBox2.BackColor = System.Drawing.Color.Transparent;
            this.GroupBox2.Controls.Add(this.chkAllowAddNew);
            this.GroupBox2.Controls.Add(this.chkAllowEdit);
            this.GroupBox2.Controls.Add(this.lblRowsFont);
            this.GroupBox2.Controls.Add(this.btnRowsFont);
            this.GroupBox2.Location = new System.Drawing.Point(8, 86);
            this.GroupBox2.Name = "GroupBox2";
            this.GroupBox2.Size = new System.Drawing.Size(424, 68);
            this.GroupBox2.TabIndex = 13;
            this.GroupBox2.Text = "Rows";
            //this.GroupBox2.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // chkAllowAddNew
            // 
            this.chkAllowAddNew.Location = new System.Drawing.Point(276, 48);
            this.chkAllowAddNew.Name = "chkAllowAddNew";
            this.chkAllowAddNew.Size = new System.Drawing.Size(132, 16);
            this.chkAllowAddNew.TabIndex = 3;
            this.chkAllowAddNew.Text = "Show \"new item\" row";
            // 
            // chkAllowEdit
            // 
            this.chkAllowEdit.Location = new System.Drawing.Point(276, 20);
            this.chkAllowEdit.Name = "chkAllowEdit";
            this.chkAllowEdit.Size = new System.Drawing.Size(124, 20);
            this.chkAllowEdit.TabIndex = 2;
            this.chkAllowEdit.Text = "Allow in-cell editing";
            // 
            // lblRowsFont
            // 
            this.lblRowsFont.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblRowsFont.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.lblRowsFont.Location = new System.Drawing.Point(100, 26);
            this.lblRowsFont.Name = "lblRowsFont";
            this.lblRowsFont.Size = new System.Drawing.Size(168, 20);
            this.lblRowsFont.TabIndex = 1;
            this.lblRowsFont.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnRowsFont
            // 
            this.btnRowsFont.Location = new System.Drawing.Point(12, 24);
            this.btnRowsFont.Name = "btnRowsFont";
            this.btnRowsFont.Size = new System.Drawing.Size(76, 24);
            this.btnRowsFont.TabIndex = 0;
            this.btnRowsFont.Text = "Font...";
            //this.btnRowsFont.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnRowsFont.Click += new System.EventHandler(this.btnRowsFont_Click);
            // 
            // GroupBox1
            // 
            this.GroupBox1.BackColor = System.Drawing.Color.Transparent;
            this.GroupBox1.Controls.Add(this.chkAutoSize);
            this.GroupBox1.Controls.Add(this.lblHeaderFont);
            this.GroupBox1.Controls.Add(this.btnHeaderFont);
            this.GroupBox1.Location = new System.Drawing.Point(8, 8);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(424, 68);
            this.GroupBox1.TabIndex = 14;
            this.GroupBox1.Text = "Column headings";
            //this.GroupBox1.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // chkAutoSize
            // 
            this.chkAutoSize.Location = new System.Drawing.Point(276, 28);
            this.chkAutoSize.Name = "chkAutoSize";
            this.chkAutoSize.Size = new System.Drawing.Size(140, 16);
            this.chkAutoSize.TabIndex = 2;
            this.chkAutoSize.Text = "Automatic column sizing";
            // 
            // lblHeaderFont
            // 
            this.lblHeaderFont.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblHeaderFont.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.lblHeaderFont.Location = new System.Drawing.Point(100, 26);
            this.lblHeaderFont.Name = "lblHeaderFont";
            this.lblHeaderFont.Size = new System.Drawing.Size(168, 20);
            this.lblHeaderFont.TabIndex = 1;
            this.lblHeaderFont.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnHeaderFont
            // 
            this.btnHeaderFont.Location = new System.Drawing.Point(12, 24);
            this.btnHeaderFont.Name = "btnHeaderFont";
            this.btnHeaderFont.Size = new System.Drawing.Size(76, 24);
            this.btnHeaderFont.TabIndex = 0;
            this.btnHeaderFont.Text = "Font...";
            //this.btnHeaderFont.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnHeaderFont.Click += new System.EventHandler(this.btnHeaderFont_Click);
            // 
            // frmFormatView
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
            this.ClientSize = new System.Drawing.Size(538, 240);
            this.Controls.Add(this.UiGroupBox1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmFormatView";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Format Table View";
            //((System.ComponentModel.ISupportInitialize)(this.UiGroupBox1)).EndInit();
            //this.UiGroupBox1.ResumeLayout(false);
            //((System.ComponentModel.ISupportInitialize)(this.GroupBox3)).EndInit();
            //this.GroupBox3.ResumeLayout(false);
            //this.GroupBox3.PerformLayout();
            //((System.ComponentModel.ISupportInitialize)(this.GroupBox2)).EndInit();
            //this.GroupBox2.ResumeLayout(false);
            //((System.ComponentModel.ISupportInitialize)(this.GroupBox1)).EndInit();
            //this.GroupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

		}

	}

} //end of root namespace