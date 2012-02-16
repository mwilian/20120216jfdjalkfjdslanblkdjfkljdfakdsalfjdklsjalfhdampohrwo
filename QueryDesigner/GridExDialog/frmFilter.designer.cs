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
	public partial class frmFilter : System.Windows.Forms.Form
	{

		//Form overrides dispose to clean up the component list.
		internal frmFilter()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmFilter));
            this.grbBackground = new Janus.Windows.EditControls.UIGroupBox();
            this.btnCancel = new Janus.Windows.EditControls.UIButton();
            this.btnOK = new Janus.Windows.EditControls.UIButton();
            this.FilterEditor1 = new Janus.Windows.FilterEditor.FilterEditor();
            ((System.ComponentModel.ISupportInitialize)(this.grbBackground)).BeginInit();
            this.grbBackground.SuspendLayout();
            this.SuspendLayout();
            // 
            // grbBackground
            // 
            this.grbBackground.BackgroundStyle = Janus.Windows.EditControls.BackgroundStyle.Panel;
            this.grbBackground.Controls.Add(this.btnCancel);
            this.grbBackground.Controls.Add(this.btnOK);
            this.grbBackground.Controls.Add(this.FilterEditor1);
            this.grbBackground.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grbBackground.FrameStyle = Janus.Windows.EditControls.FrameStyle.None;
            this.grbBackground.Location = new System.Drawing.Point(0, 0);
            this.grbBackground.Name = "grbBackground";
            this.grbBackground.Size = new System.Drawing.Size(483, 341);
            this.grbBackground.TabIndex = 1;
            this.grbBackground.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(247, 307);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(84, 24);
            this.btnCancel.TabIndex = 16;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(155, 307);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(84, 24);
            this.btnOK.TabIndex = 15;
            this.btnOK.Text = "OK";
            this.btnOK.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            // 
            // FilterEditor1
            // 
            this.FilterEditor1.AllowFilterByFieldValue = true;
            this.FilterEditor1.BackColor = System.Drawing.Color.Transparent;
            this.FilterEditor1.Dock = System.Windows.Forms.DockStyle.Top;
            this.FilterEditor1.InnerAreaStyle = Janus.Windows.UI.Dock.PanelInnerAreaStyle.UseFormatStyle;
            this.FilterEditor1.Location = new System.Drawing.Point(0, 0);
            this.FilterEditor1.MinSize = new System.Drawing.Size(0, 0);
            this.FilterEditor1.Name = "FilterEditor1";
            this.FilterEditor1.ScrollMode = Janus.Windows.UI.Dock.ScrollMode.Both;
            this.FilterEditor1.ScrollStep = 15;
            this.FilterEditor1.Size = new System.Drawing.Size(483, 301);
            // 
            // frmFilter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(483, 341);
            this.Controls.Add(this.grbBackground);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmFilter";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Filter Conditions";
            ((System.ComponentModel.ISupportInitialize)(this.grbBackground)).EndInit();
            this.grbBackground.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		internal Janus.Windows.EditControls.UIGroupBox grbBackground;
		internal Janus.Windows.FilterEditor.FilterEditor FilterEditor1;
		internal Janus.Windows.EditControls.UIButton btnCancel;
		internal Janus.Windows.EditControls.UIButton btnOK;
	}

} //end of root namespace