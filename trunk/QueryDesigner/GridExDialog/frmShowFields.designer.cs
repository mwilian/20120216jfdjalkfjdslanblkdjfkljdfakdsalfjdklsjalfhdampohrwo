using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;

using Janus.Windows.GridEX;

namespace dCube
{
	public partial class frmShowFields : System.Windows.Forms.Form
	{

		//Form overrides dispose to clean up the component list.
		internal frmShowFields()
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
        internal Button btnCancel;
        internal Button btnOk;
        internal Button btnDown;
        internal Button btnUp;
        internal Button btnRemove;
        internal Button btnAdd;
		internal System.Windows.Forms.ListBox lbVisible;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.ListBox lbAvail;
        internal GroupBox grbBackground;
		private void InitializeComponent()
		{
            this.grbBackground = new GroupBox();
            this.btnCancel = new Button();
            this.btnOk = new Button();
            this.btnDown = new Button();
            this.btnUp = new Button();
            this.btnRemove = new Button();
            this.btnAdd = new Button();
            this.lbVisible = new System.Windows.Forms.ListBox();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.lbAvail = new System.Windows.Forms.ListBox();
            //((System.ComponentModel.ISupportInitialize)(this.grbBackground)).BeginInit();
            //this.grbBackground.SuspendLayout();
            this.SuspendLayout();
            // 
            // grbBackground
            // 
            //this.grbBackground.BackgroundStyle = Janus.Windows.EditControls.BackgroundStyle.Panel;
            this.grbBackground.Controls.Add(this.btnCancel);
            this.grbBackground.Controls.Add(this.btnOk);
            this.grbBackground.Controls.Add(this.btnDown);
            this.grbBackground.Controls.Add(this.btnUp);
            this.grbBackground.Controls.Add(this.btnRemove);
            this.grbBackground.Controls.Add(this.btnAdd);
            this.grbBackground.Controls.Add(this.lbVisible);
            this.grbBackground.Controls.Add(this.Label2);
            this.grbBackground.Controls.Add(this.Label1);
            this.grbBackground.Controls.Add(this.lbAvail);
            this.grbBackground.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.grbBackground.FrameStyle = Janus.Windows.EditControls.FrameStyle.None;
            this.grbBackground.Location = new System.Drawing.Point(0, 0);
            this.grbBackground.Name = "grbBackground";
            this.grbBackground.Size = new System.Drawing.Size(482, 240);
            this.grbBackground.TabIndex = 0;
            //this.grbBackground.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(393, 208);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(80, 24);
            this.btnCancel.TabIndex = 25;
            this.btnCancel.Text = "Cancel";
            //this.btnCancel.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOk
            // 
            this.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOk.Location = new System.Drawing.Point(305, 208);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(80, 24);
            this.btnOk.TabIndex = 24;
            this.btnOk.Text = "OK";
            //this.btnOk.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnOk.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnDown
            // 
            this.btnDown.Location = new System.Drawing.Point(201, 124);
            this.btnDown.Name = "btnDown";
            this.btnDown.Size = new System.Drawing.Size(83, 24);
            this.btnDown.TabIndex = 23;
            this.btnDown.Text = "Move Down";
            //this.btnDown.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
            // 
            // btnUp
            // 
            this.btnUp.Location = new System.Drawing.Point(201, 92);
            this.btnUp.Name = "btnUp";
            this.btnUp.Size = new System.Drawing.Size(83, 24);
            this.btnUp.TabIndex = 22;
            this.btnUp.Text = "Move Up";
            //this.btnUp.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(201, 60);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(83, 24);
            this.btnRemove.TabIndex = 21;
            this.btnRemove.Text = "<- Remove";
            //this.btnRemove.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(201, 28);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(83, 24);
            this.btnAdd.TabIndex = 20;
            this.btnAdd.Text = "Add ->";
            //this.btnAdd.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // lbVisible
            // 
            this.lbVisible.Location = new System.Drawing.Point(289, 28);
            this.lbVisible.Name = "lbVisible";
            this.lbVisible.Size = new System.Drawing.Size(184, 173);
            this.lbVisible.TabIndex = 19;
            this.lbVisible.SelectedIndexChanged += new System.EventHandler(this.lbVisible_SelectedIndexChanged);
            this.lbVisible.DoubleClick += new System.EventHandler(this.lbVisible_DoubleClick);
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.BackColor = System.Drawing.Color.Transparent;
            this.Label2.Location = new System.Drawing.Point(289, 8);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(155, 13);
            this.Label2.TabIndex = 18;
            this.Label2.Text = "Show these fields in this order:";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.BackColor = System.Drawing.Color.Transparent;
            this.Label1.Location = new System.Drawing.Point(9, 8);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(84, 13);
            this.Label1.TabIndex = 17;
            this.Label1.Text = "Available _Fields:";
            // 
            // lbAvail
            // 
            this.lbAvail.Location = new System.Drawing.Point(9, 28);
            this.lbAvail.Name = "lbAvail";
            this.lbAvail.Size = new System.Drawing.Size(184, 173);
            this.lbAvail.TabIndex = 16;
            this.lbAvail.DoubleClick += new System.EventHandler(this.lbAvail_DoubleClick);
            // 
            // frmShowFields
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
            this.ClientSize = new System.Drawing.Size(482, 240);
            this.Controls.Add(this.grbBackground);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmShowFields";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Show _Fields";
            //((System.ComponentModel.ISupportInitialize)(this.grbBackground)).EndInit();
            //this.grbBackground.ResumeLayout(false);
            this.grbBackground.PerformLayout();
            this.ResumeLayout(false);

        }

	}

} //end of root namespace