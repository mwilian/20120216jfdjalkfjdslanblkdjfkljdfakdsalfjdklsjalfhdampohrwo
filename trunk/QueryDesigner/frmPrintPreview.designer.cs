using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;

namespace QueryDesigner
{
	[Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
	public partial class frmPrintPreview : System.Windows.Forms.Form
	{


		//Form overrides dispose to clean up the component list.
		internal frmPrintPreview()
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
		internal System.Windows.Forms.ImageList icons;
		internal ToolStrip UiCommandBar1;
		internal ToolStripButton cmdMoveUp;
		internal ToolStripButton cmdMoveDown;
		internal ToolStripButton cmdZoom100;
		internal ToolStripButton cmdOnePage;
		internal ToolStripButton cmdTwoPages;
		internal ToolStripButton cmdPageSetup;
		internal ToolStripButton cmdPrint;
		internal ToolStripButton cmdClose;
		internal ToolStripButton cmdMoveUp1;
		internal ToolStripButton cmdMoveDown1;
		internal ToolStripButton cmdSeparator1;
		internal ToolStripButton cmdZoom1001;
		internal ToolStripButton cmdOnePage1;
		internal ToolStripButton cmdTwoPages1;
		internal ToolStripButton cmdSeparator2;
		internal ToolStripButton cmdPageSetup1;
		internal ToolStripButton cmdSeparator3;
		internal ToolStripButton cmdPrint1;
		internal ToolStripButton cmdSeparator4;
		internal ToolStripButton cmdClose1;
        private PrintPreviewControl PrintPreviewControl1;
        private PageSetupDialog PageSetupDialog1;
        private ToolStrip toolStrip1;
        private ToolStripButton toolStripButton1;
        private ToolStripSeparator toolStripSeparator1;
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPrintPreview));
            this.cmdMoveUp = new ToolStripButton("cmdMoveUp");
            this.cmdMoveDown = new ToolStripButton("cmdMoveDown");
            this.cmdZoom100 = new ToolStripButton("cmdZoom100");
            this.cmdOnePage = new ToolStripButton("cmdOnePage");
            this.cmdTwoPages = new ToolStripButton("cmdTwoPages");
            this.cmdPageSetup = new ToolStripButton("cmdPageSetup");
            this.cmdPrint = new ToolStripButton("cmdPrint");
            this.cmdClose = new ToolStripButton("cmdClose");
            this.cmdMoveUp1 = new ToolStripButton("cmdMoveUp");
            this.cmdMoveDown1 = new ToolStripButton("cmdMoveDown");
            this.cmdSeparator1 = new ToolStripButton("Separator");
            this.cmdZoom1001 = new ToolStripButton("cmdZoom100");
            this.cmdOnePage1 = new ToolStripButton("cmdOnePage");
            this.cmdTwoPages1 = new ToolStripButton("cmdTwoPages");
            this.cmdSeparator2 = new ToolStripButton("Separator");
            this.cmdPageSetup1 = new ToolStripButton("cmdPageSetup");
            this.cmdSeparator3 = new ToolStripButton("Separator");
            this.cmdPrint1 = new ToolStripButton("cmdPrint");
            this.cmdSeparator4 = new ToolStripButton("Separator");
            this.cmdClose1 = new ToolStripButton("cmdClose");
            this.icons = new System.Windows.Forms.ImageList(this.components);
            this.PrintPreviewControl1 = new System.Windows.Forms.PrintPreviewControl();
            this.PageSetupDialog1 = new System.Windows.Forms.PageSetupDialog();
            this.UiCommandBar1 = new ToolStrip();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmdMoveUp
            // 
            this.cmdMoveUp.ImageIndex = 0;
            this.cmdMoveUp.Name = "cmdMoveUp";
            this.cmdMoveUp.Text = "Page Up";
            this.cmdMoveUp.ToolTipText = "Page Up";
            // 
            // cmdMoveDown
            // 
            this.cmdMoveDown.ImageIndex = 1;
            this.cmdMoveDown.Name = "cmdMoveDown";
            this.cmdMoveDown.Text = "Page Down";
            this.cmdMoveDown.ToolTipText = "Page Down";
            // 
            // cmdZoom100
            // 
            this.cmdZoom100.ImageIndex = 3;
            this.cmdZoom100.Name = "cmdZoom100";
            this.cmdZoom100.Text = "Actual Size";
            this.cmdZoom100.ToolTipText = "Actual Size";
            // 
            // cmdOnePage
            // 
            this.cmdOnePage.ImageIndex = 2;
            this.cmdOnePage.Name = "cmdOnePage";
            this.cmdOnePage.Text = "One Page";
            this.cmdOnePage.ToolTipText = "One Page";
            // 
            // cmdTwoPages
            // 
            this.cmdTwoPages.ImageIndex = 4;
            this.cmdTwoPages.Name = "cmdTwoPages";
            this.cmdTwoPages.Text = "Two Pages";
            this.cmdTwoPages.ToolTipText = "Two Pages";
            // 
            // cmdPageSetup
            // 
            this.cmdPageSetup.ImageIndex = 6;
            this.cmdPageSetup.Name = "cmdPageSetup";
            this.cmdPageSetup.Text = "Page Setup...";
            this.cmdPageSetup.ToolTipText = "Page Setup";
            // 
            // cmdPrint
            // 
            this.cmdPrint.ImageIndex = 5;
            this.cmdPrint.Name = "cmdPrint";
            this.cmdPrint.Text = "Print";
            this.cmdPrint.ToolTipText = "Print";
            // 
            // cmdClose
            // 
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.Text = "Close";
            this.cmdClose.ToolTipText = "Close Preview";
            // 
            // cmdMoveUp1
            // 
            this.cmdMoveUp1.Name = "cmdMoveUp1";
            // 
            // cmdMoveDown1
            // 
            this.cmdMoveDown1.Name = "cmdMoveDown1";
            // 
            // cmdSeparator1
            // 
            this.cmdSeparator1.Name = "cmdSeparator1";
            // 
            // cmdZoom1001
            // 
            this.cmdZoom1001.Name = "cmdZoom1001";
            // 
            // cmdOnePage1
            // 
            this.cmdOnePage1.Name = "cmdOnePage1";
            // 
            // cmdTwoPages1
            // 
            this.cmdTwoPages1.Name = "cmdTwoPages1";
            // 
            // cmdSeparator2
            // 
            this.cmdSeparator2.Name = "cmdSeparator2";
            // 
            // cmdPageSetup1
            // 
            this.cmdPageSetup1.Name = "cmdPageSetup1";
            // 
            // cmdSeparator3
            // 
            this.cmdSeparator3.Name = "cmdSeparator3";
            // 
            // cmdPrint1
            // 
            this.cmdPrint1.Name = "cmdPrint1";
            // 
            // cmdSeparator4
            // 
            this.cmdSeparator4.Name = "cmdSeparator4";
            // 
            // cmdClose1
            // 
            this.cmdClose1.Name = "cmdClose1";
            // 
            // icons
            // 
            this.icons.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("icons.ImageStream")));
            this.icons.TransparentColor = System.Drawing.Color.Transparent;
            this.icons.Images.SetKeyName(0, "");
            this.icons.Images.SetKeyName(1, "");
            this.icons.Images.SetKeyName(2, "");
            this.icons.Images.SetKeyName(3, "");
            this.icons.Images.SetKeyName(4, "");
            this.icons.Images.SetKeyName(5, "");
            this.icons.Images.SetKeyName(6, "");
            // 
            // PrintPreviewControl1
            // 
            this.PrintPreviewControl1.AutoZoom = false;
            this.PrintPreviewControl1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.PrintPreviewControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.PrintPreviewControl1.Location = new System.Drawing.Point(0, 28);
            this.PrintPreviewControl1.Name = "PrintPreviewControl1";
            this.PrintPreviewControl1.Size = new System.Drawing.Size(708, 390);
            this.PrintPreviewControl1.TabIndex = 1;
            this.PrintPreviewControl1.UseAntiAlias = true;
            this.PrintPreviewControl1.Zoom = 1;
            this.PrintPreviewControl1.StartPageChanged += new System.EventHandler(this.PrintPreviewControl1_StartPageChanged);
           
            // 
            // UiCommandBar1
            // 
            this.UiCommandBar1.Items.AddRange(new ToolStripButton[] {
            this.cmdMoveUp1,
            this.cmdMoveDown1,
            this.cmdSeparator1,
            this.cmdZoom1001,
            this.cmdOnePage1,
            this.cmdTwoPages1,
            this.cmdSeparator2,
            this.cmdPageSetup1,
            this.cmdSeparator3,
            this.cmdPrint1,
            this.cmdSeparator4,
            this.cmdClose1});
            this.UiCommandBar1.Location = new System.Drawing.Point(0, 0);
            this.UiCommandBar1.Name = "UiCommandBar1";
            this.UiCommandBar1.Size = new System.Drawing.Size(333, 28);
            this.UiCommandBar1.Text = "Print Preview";           
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton1,
            this.toolStripSeparator1});
            this.toolStrip1.Location = new System.Drawing.Point(0, 28);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(708, 25);
            this.toolStrip1.TabIndex = 2;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(23, 22);
            this.toolStripButton1.Text = "toolStripButton1";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // frmPrintPreview
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
            this.ClientSize = new System.Drawing.Size(708, 418);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.UiCommandBar1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmPrintPreview";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "frmPrintPreview";
            this.Load += new System.EventHandler(this.frmPrintPreview_Load);
            ((System.ComponentModel.ISupportInitialize)(this.UiCommandBar1)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

        

	}

} //end of root namespace