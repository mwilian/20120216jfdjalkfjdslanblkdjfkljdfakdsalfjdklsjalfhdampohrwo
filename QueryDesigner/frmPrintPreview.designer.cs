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
		internal Janus.Windows.UI.CommandBars.UICommandManager printPreviewCommands;
		internal Janus.Windows.UI.CommandBars.UICommandBar UiCommandBar1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdMoveUp;
		internal Janus.Windows.UI.CommandBars.UICommand cmdMoveDown;
		internal Janus.Windows.UI.CommandBars.UICommand cmdZoom100;
		internal Janus.Windows.UI.CommandBars.UICommand cmdOnePage;
		internal Janus.Windows.UI.CommandBars.UICommand cmdTwoPages;
		internal Janus.Windows.UI.CommandBars.UICommand cmdPageSetup;
		internal Janus.Windows.UI.CommandBars.UICommand cmdPrint;
		internal Janus.Windows.UI.CommandBars.UICommand cmdClose;
		internal Janus.Windows.UI.CommandBars.UICommand cmdMoveUp1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdMoveDown1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdSeparator1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdZoom1001;
		internal Janus.Windows.UI.CommandBars.UICommand cmdOnePage1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdTwoPages1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdSeparator2;
		internal Janus.Windows.UI.CommandBars.UICommand cmdPageSetup1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdSeparator3;
		internal Janus.Windows.UI.CommandBars.UICommand cmdPrint1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdSeparator4;
		internal Janus.Windows.UI.CommandBars.UICommand cmdClose1;
		internal Janus.Windows.UI.CommandBars.UIRebar TopRebar1;
		internal Janus.Windows.UI.CommandBars.UIRebar BottomRebar1;
		internal Janus.Windows.UI.CommandBars.UIRebar LeftRebar1;
		internal Janus.Windows.UI.CommandBars.UIRebar RightRebar1;
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPrintPreview));
            this.cmdMoveUp = new Janus.Windows.UI.CommandBars.UICommand("cmdMoveUp");
            this.cmdMoveDown = new Janus.Windows.UI.CommandBars.UICommand("cmdMoveDown");
            this.cmdZoom100 = new Janus.Windows.UI.CommandBars.UICommand("cmdZoom100");
            this.cmdOnePage = new Janus.Windows.UI.CommandBars.UICommand("cmdOnePage");
            this.cmdTwoPages = new Janus.Windows.UI.CommandBars.UICommand("cmdTwoPages");
            this.cmdPageSetup = new Janus.Windows.UI.CommandBars.UICommand("cmdPageSetup");
            this.cmdPrint = new Janus.Windows.UI.CommandBars.UICommand("cmdPrint");
            this.cmdClose = new Janus.Windows.UI.CommandBars.UICommand("cmdClose");
            this.cmdMoveUp1 = new Janus.Windows.UI.CommandBars.UICommand("cmdMoveUp");
            this.cmdMoveDown1 = new Janus.Windows.UI.CommandBars.UICommand("cmdMoveDown");
            this.cmdSeparator1 = new Janus.Windows.UI.CommandBars.UICommand("Separator");
            this.cmdZoom1001 = new Janus.Windows.UI.CommandBars.UICommand("cmdZoom100");
            this.cmdOnePage1 = new Janus.Windows.UI.CommandBars.UICommand("cmdOnePage");
            this.cmdTwoPages1 = new Janus.Windows.UI.CommandBars.UICommand("cmdTwoPages");
            this.cmdSeparator2 = new Janus.Windows.UI.CommandBars.UICommand("Separator");
            this.cmdPageSetup1 = new Janus.Windows.UI.CommandBars.UICommand("cmdPageSetup");
            this.cmdSeparator3 = new Janus.Windows.UI.CommandBars.UICommand("Separator");
            this.cmdPrint1 = new Janus.Windows.UI.CommandBars.UICommand("cmdPrint");
            this.cmdSeparator4 = new Janus.Windows.UI.CommandBars.UICommand("Separator");
            this.cmdClose1 = new Janus.Windows.UI.CommandBars.UICommand("cmdClose");
            this.icons = new System.Windows.Forms.ImageList(this.components);
            this.PrintPreviewControl1 = new System.Windows.Forms.PrintPreviewControl();
            this.PageSetupDialog1 = new System.Windows.Forms.PageSetupDialog();
            this.printPreviewCommands = new Janus.Windows.UI.CommandBars.UICommandManager(this.components);
            this.BottomRebar1 = new Janus.Windows.UI.CommandBars.UIRebar();
            this.UiCommandBar1 = new Janus.Windows.UI.CommandBars.UICommandBar();
            this.LeftRebar1 = new Janus.Windows.UI.CommandBars.UIRebar();
            this.RightRebar1 = new Janus.Windows.UI.CommandBars.UIRebar();
            this.TopRebar1 = new Janus.Windows.UI.CommandBars.UIRebar();
            ((System.ComponentModel.ISupportInitialize)(this.printPreviewCommands)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.BottomRebar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.UiCommandBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LeftRebar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RightRebar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TopRebar1)).BeginInit();
            this.TopRebar1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmdMoveUp
            // 
            this.cmdMoveUp.CommandStyle = Janus.Windows.UI.CommandBars.CommandStyle.Image;
            this.cmdMoveUp.Enabled = Janus.Windows.UI.InheritableBoolean.False;
            this.cmdMoveUp.ImageIndex = 0;
            this.cmdMoveUp.Key = "cmdMoveUp";
            this.cmdMoveUp.Name = "cmdMoveUp";
            this.cmdMoveUp.Text = "Page Up";
            this.cmdMoveUp.ToolTipText = "Page Up";
            // 
            // cmdMoveDown
            // 
            this.cmdMoveDown.CommandStyle = Janus.Windows.UI.CommandBars.CommandStyle.Image;
            this.cmdMoveDown.ImageIndex = 1;
            this.cmdMoveDown.Key = "cmdMoveDown";
            this.cmdMoveDown.Name = "cmdMoveDown";
            this.cmdMoveDown.Text = "Page Down";
            this.cmdMoveDown.ToolTipText = "Page Down";
            // 
            // cmdZoom100
            // 
            this.cmdZoom100.CommandStyle = Janus.Windows.UI.CommandBars.CommandStyle.Image;
            this.cmdZoom100.CommandType = Janus.Windows.UI.CommandBars.CommandType.ToggleButton;
            this.cmdZoom100.ImageIndex = 3;
            this.cmdZoom100.Key = "cmdZoom100";
            this.cmdZoom100.Name = "cmdZoom100";
            this.cmdZoom100.Text = "Actual Size";
            this.cmdZoom100.ToolTipText = "Actual Size";
            // 
            // cmdOnePage
            // 
            this.cmdOnePage.Checked = Janus.Windows.UI.InheritableBoolean.True;
            this.cmdOnePage.CommandStyle = Janus.Windows.UI.CommandBars.CommandStyle.Image;
            this.cmdOnePage.CommandType = Janus.Windows.UI.CommandBars.CommandType.ToggleButton;
            this.cmdOnePage.ImageIndex = 2;
            this.cmdOnePage.Key = "cmdOnePage";
            this.cmdOnePage.Name = "cmdOnePage";
            this.cmdOnePage.Text = "One Page";
            this.cmdOnePage.ToolTipText = "One Page";
            // 
            // cmdTwoPages
            // 
            this.cmdTwoPages.CommandStyle = Janus.Windows.UI.CommandBars.CommandStyle.Image;
            this.cmdTwoPages.CommandType = Janus.Windows.UI.CommandBars.CommandType.ToggleButton;
            this.cmdTwoPages.ImageIndex = 4;
            this.cmdTwoPages.Key = "cmdTwoPages";
            this.cmdTwoPages.Name = "cmdTwoPages";
            this.cmdTwoPages.Text = "Two Pages";
            this.cmdTwoPages.ToolTipText = "Two Pages";
            // 
            // cmdPageSetup
            // 
            this.cmdPageSetup.ImageIndex = 6;
            this.cmdPageSetup.Key = "cmdPageSetup";
            this.cmdPageSetup.Name = "cmdPageSetup";
            this.cmdPageSetup.Text = "Page Setup...";
            this.cmdPageSetup.ToolTipText = "Page Setup";
            // 
            // cmdPrint
            // 
            this.cmdPrint.ImageIndex = 5;
            this.cmdPrint.Key = "cmdPrint";
            this.cmdPrint.Name = "cmdPrint";
            this.cmdPrint.Text = "Print";
            this.cmdPrint.ToolTipText = "Print";
            // 
            // cmdClose
            // 
            this.cmdClose.Key = "cmdClose";
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.Text = "Close";
            this.cmdClose.ToolTipText = "Close Preview";
            // 
            // cmdMoveUp1
            // 
            this.cmdMoveUp1.Key = "cmdMoveUp";
            this.cmdMoveUp1.Name = "cmdMoveUp1";
            this.cmdMoveUp1.Visible = Janus.Windows.UI.InheritableBoolean.False;
            // 
            // cmdMoveDown1
            // 
            this.cmdMoveDown1.Key = "cmdMoveDown";
            this.cmdMoveDown1.Name = "cmdMoveDown1";
            // 
            // cmdSeparator1
            // 
            this.cmdSeparator1.CommandType = Janus.Windows.UI.CommandBars.CommandType.Separator;
            this.cmdSeparator1.Key = "Separator";
            this.cmdSeparator1.Name = "cmdSeparator1";
            // 
            // cmdZoom1001
            // 
            this.cmdZoom1001.Key = "cmdZoom100";
            this.cmdZoom1001.Name = "cmdZoom1001";
            // 
            // cmdOnePage1
            // 
            this.cmdOnePage1.Key = "cmdOnePage";
            this.cmdOnePage1.Name = "cmdOnePage1";
            // 
            // cmdTwoPages1
            // 
            this.cmdTwoPages1.Key = "cmdTwoPages";
            this.cmdTwoPages1.Name = "cmdTwoPages1";
            // 
            // cmdSeparator2
            // 
            this.cmdSeparator2.CommandType = Janus.Windows.UI.CommandBars.CommandType.Separator;
            this.cmdSeparator2.Key = "Separator";
            this.cmdSeparator2.Name = "cmdSeparator2";
            // 
            // cmdPageSetup1
            // 
            this.cmdPageSetup1.Key = "cmdPageSetup";
            this.cmdPageSetup1.Name = "cmdPageSetup1";
            // 
            // cmdSeparator3
            // 
            this.cmdSeparator3.CommandType = Janus.Windows.UI.CommandBars.CommandType.Separator;
            this.cmdSeparator3.Key = "Separator";
            this.cmdSeparator3.Name = "cmdSeparator3";
            // 
            // cmdPrint1
            // 
            this.cmdPrint1.Key = "cmdPrint";
            this.cmdPrint1.Name = "cmdPrint1";
            // 
            // cmdSeparator4
            // 
            this.cmdSeparator4.CommandType = Janus.Windows.UI.CommandBars.CommandType.Separator;
            this.cmdSeparator4.Key = "Separator";
            this.cmdSeparator4.Name = "cmdSeparator4";
            // 
            // cmdClose1
            // 
            this.cmdClose1.Key = "cmdClose";
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
            // printPreviewCommands
            // 
            this.printPreviewCommands.AllowCustomize = Janus.Windows.UI.InheritableBoolean.True;
            this.printPreviewCommands.AlwaysShowFullMenus = true;
            this.printPreviewCommands.BottomRebar = this.BottomRebar1;
            this.printPreviewCommands.CommandBars.AddRange(new Janus.Windows.UI.CommandBars.UICommandBar[] {
            this.UiCommandBar1});
            this.printPreviewCommands.Commands.AddRange(new Janus.Windows.UI.CommandBars.UICommand[] {
            this.cmdMoveUp,
            this.cmdMoveDown,
            this.cmdZoom100,
            this.cmdOnePage,
            this.cmdTwoPages,
            this.cmdPageSetup,
            this.cmdPrint,
            this.cmdClose});
            this.printPreviewCommands.ContainerControl = this;
            this.printPreviewCommands.Id = new System.Guid("e2e0f92e-194e-4949-9680-5d6b788c72e3");
            this.printPreviewCommands.ImageList = this.icons;
            this.printPreviewCommands.LeftRebar = this.LeftRebar1;
            this.printPreviewCommands.RightRebar = this.RightRebar1;
            this.printPreviewCommands.TopRebar = this.TopRebar1;
            this.printPreviewCommands.VisualStyle = Janus.Windows.UI.VisualStyle.Office2007;
            this.printPreviewCommands.CommandClick += new Janus.Windows.UI.CommandBars.CommandEventHandler(this.printPreviewCommands_CommandClick);
            // 
            // BottomRebar1
            // 
            this.BottomRebar1.CommandManager = this.printPreviewCommands;
            this.BottomRebar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.BottomRebar1.Location = new System.Drawing.Point(0, 0);
            this.BottomRebar1.Name = "BottomRebar1";
            this.BottomRebar1.Size = new System.Drawing.Size(0, 0);
            // 
            // UiCommandBar1
            // 
            this.UiCommandBar1.CommandManager = this.printPreviewCommands;
            this.UiCommandBar1.Commands.AddRange(new Janus.Windows.UI.CommandBars.UICommand[] {
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
            this.UiCommandBar1.Key = "PrintPreview";
            this.UiCommandBar1.Location = new System.Drawing.Point(0, 0);
            this.UiCommandBar1.Name = "UiCommandBar1";
            this.UiCommandBar1.RowIndex = 0;
            this.UiCommandBar1.Size = new System.Drawing.Size(333, 28);
            this.UiCommandBar1.Text = "Print Preview";
            // 
            // LeftRebar1
            // 
            this.LeftRebar1.CommandManager = this.printPreviewCommands;
            this.LeftRebar1.Dock = System.Windows.Forms.DockStyle.Left;
            this.LeftRebar1.Location = new System.Drawing.Point(0, 0);
            this.LeftRebar1.Name = "LeftRebar1";
            this.LeftRebar1.Size = new System.Drawing.Size(0, 0);
            // 
            // RightRebar1
            // 
            this.RightRebar1.CommandManager = this.printPreviewCommands;
            this.RightRebar1.Dock = System.Windows.Forms.DockStyle.Right;
            this.RightRebar1.Location = new System.Drawing.Point(0, 0);
            this.RightRebar1.Name = "RightRebar1";
            this.RightRebar1.Size = new System.Drawing.Size(0, 0);
            // 
            // TopRebar1
            // 
            this.TopRebar1.CommandBars.AddRange(new Janus.Windows.UI.CommandBars.UICommandBar[] {
            this.UiCommandBar1});
            this.TopRebar1.CommandManager = this.printPreviewCommands;
            this.TopRebar1.Controls.Add(this.UiCommandBar1);
            this.TopRebar1.Dock = System.Windows.Forms.DockStyle.Top;
            this.TopRebar1.Location = new System.Drawing.Point(0, 0);
            this.TopRebar1.Name = "TopRebar1";
            this.TopRebar1.Size = new System.Drawing.Size(708, 28);
            // 
            // frmPrintPreview
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
            this.ClientSize = new System.Drawing.Size(708, 418);
            this.Controls.Add(this.PrintPreviewControl1);
            this.Controls.Add(this.TopRebar1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmPrintPreview";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "frmPrintPreview";
            this.Load += new System.EventHandler(this.frmPrintPreview_Load);
            ((System.ComponentModel.ISupportInitialize)(this.printPreviewCommands)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.BottomRebar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.UiCommandBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LeftRebar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RightRebar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TopRebar1)).EndInit();
            this.TopRebar1.ResumeLayout(false);
            this.ResumeLayout(false);

		}

        private PrintPreviewControl PrintPreviewControl1;
        private PageSetupDialog PageSetupDialog1;

	}

} //end of root namespace