using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;

namespace TVCDesigner
{
	public partial class MainForm : System.Windows.Forms.Form
	{
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.StatusBar statusBar1;
        private System.Windows.Forms.MainMenu MainMenu;
        private System.Windows.Forms.MenuItem File;
        private System.Windows.Forms.MenuItem miExit;
        private System.Windows.Forms.MenuItem miAlwaysOnTop;
        private System.Windows.Forms.MenuItem menuItem3;
        private System.Windows.Forms.MenuItem menuItem4;
        private System.Windows.Forms.MenuItem menuItem5;
        private System.Windows.Forms.MenuItem menuItem6;
        private System.Windows.Forms.MenuItem menuItem7;
        private System.Windows.Forms.MenuItem miOpacity;
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ImageList imageListTree;
        private System.Windows.Forms.TreeView tvFields;
        private System.Windows.Forms.MenuItem menuItem1;
        private System.Windows.Forms.MenuItem menuItem2;
        private System.Windows.Forms.MenuItem miUseColumnCaptions;
        private System.Windows.Forms.MenuItem miOpen;
        private System.Windows.Forms.OpenFileDialog openXls;
        private System.Windows.Forms.MenuItem miCopy;
        private System.Windows.Forms.MenuItem menuItem9;


		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.statusBar1 = new System.Windows.Forms.StatusBar();
            this.tvFields = new System.Windows.Forms.TreeView();
            this.imageListTree = new System.Windows.Forms.ImageList(this.components);
            this.MainMenu = new System.Windows.Forms.MainMenu(this.components);
            this.File = new System.Windows.Forms.MenuItem();
            this.miOpen = new System.Windows.Forms.MenuItem();
            this.menuItem9 = new System.Windows.Forms.MenuItem();
            this.miExit = new System.Windows.Forms.MenuItem();
            this.menuItem2 = new System.Windows.Forms.MenuItem();
            this.miUseColumnCaptions = new System.Windows.Forms.MenuItem();
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.miAlwaysOnTop = new System.Windows.Forms.MenuItem();
            this.miOpacity = new System.Windows.Forms.MenuItem();
            this.menuItem3 = new System.Windows.Forms.MenuItem();
            this.menuItem4 = new System.Windows.Forms.MenuItem();
            this.menuItem5 = new System.Windows.Forms.MenuItem();
            this.menuItem6 = new System.Windows.Forms.MenuItem();
            this.menuItem7 = new System.Windows.Forms.MenuItem();
            this.miCopy = new System.Windows.Forms.MenuItem();
            this.openXls = new System.Windows.Forms.OpenFileDialog();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Gray;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.ForeColor = System.Drawing.Color.White;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(239, 21);
            this.panel1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label1.Location = new System.Drawing.Point(8, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(224, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Drag and drop fields to Excel.";
            // 
            // statusBar1
            // 
            this.statusBar1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.statusBar1.Location = new System.Drawing.Point(0, 233);
            this.statusBar1.Name = "statusBar1";
            this.statusBar1.Size = new System.Drawing.Size(239, 16);
            this.statusBar1.TabIndex = 1;
            this.statusBar1.Text = "Alt-click to drag captions";
            // 
            // tvFields
            // 
            this.tvFields.AllowDrop = true;
            this.tvFields.BackColor = System.Drawing.SystemColors.Window;
            this.tvFields.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvFields.HideSelection = false;
            this.tvFields.ImageIndex = 0;
            this.tvFields.ImageList = this.imageListTree;
            this.tvFields.Indent = 19;
            this.tvFields.ItemHeight = 16;
            this.tvFields.Location = new System.Drawing.Point(0, 21);
            this.tvFields.Name = "tvFields";
            this.tvFields.SelectedImageIndex = 0;
            this.tvFields.Size = new System.Drawing.Size(239, 212);
            this.tvFields.TabIndex = 2;
            this.tvFields.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.tvFields_ItemDrag);
            this.tvFields.DragDrop += new System.Windows.Forms.DragEventHandler(this.tvFields_DragDrop);
            this.tvFields.DragOver += new System.Windows.Forms.DragEventHandler(this.tvFields_DragOver);
            // 
            // imageListTree
            // 
            this.imageListTree.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListTree.ImageStream")));
            this.imageListTree.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListTree.Images.SetKeyName(0, "");
            this.imageListTree.Images.SetKeyName(1, "");
            this.imageListTree.Images.SetKeyName(2, "");
            this.imageListTree.Images.SetKeyName(3, "");
            this.imageListTree.Images.SetKeyName(4, "");
            this.imageListTree.Images.SetKeyName(5, "");
            this.imageListTree.Images.SetKeyName(6, "");
            this.imageListTree.Images.SetKeyName(7, "");
            this.imageListTree.Images.SetKeyName(8, "");
            this.imageListTree.Images.SetKeyName(9, "");
            this.imageListTree.Images.SetKeyName(10, "");
            this.imageListTree.Images.SetKeyName(11, "");
            this.imageListTree.Images.SetKeyName(12, "");
            this.imageListTree.Images.SetKeyName(13, "");
            this.imageListTree.Images.SetKeyName(14, "");
            this.imageListTree.Images.SetKeyName(15, "");
            this.imageListTree.Images.SetKeyName(16, "");
            // 
            // MainMenu
            // 
            this.MainMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.File,
            this.menuItem2,
            this.menuItem1,
            this.miCopy});
            // 
            // File
            // 
            this.File.Index = 0;
            this.File.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.miOpen,
            this.menuItem9,
            this.miExit});
            this.File.Text = "File";
            this.File.Click += new System.EventHandler(this.File_Click);
            // 
            // miOpen
            // 
            this.miOpen.Index = 0;
            this.miOpen.Text = "Open...";
            this.miOpen.Click += new System.EventHandler(this.miOpen_Click);
            // 
            // menuItem9
            // 
            this.menuItem9.Index = 1;
            this.menuItem9.Text = "-";
            // 
            // miExit
            // 
            this.miExit.Index = 2;
            this.miExit.Text = "Exit";
            this.miExit.Click += new System.EventHandler(this.miExit_Click);
            // 
            // menuItem2
            // 
            this.menuItem2.Index = 1;
            this.menuItem2.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.miUseColumnCaptions});
            this.menuItem2.Text = "Options";
            // 
            // miUseColumnCaptions
            // 
            this.miUseColumnCaptions.Checked = true;
            this.miUseColumnCaptions.Index = 0;
            this.miUseColumnCaptions.Text = "Use column captions";
            this.miUseColumnCaptions.Click += new System.EventHandler(this.miUseColumnCaptions_Click);
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 2;
            this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.miAlwaysOnTop,
            this.miOpacity});
            this.menuItem1.Text = "View";
            // 
            // miAlwaysOnTop
            // 
            this.miAlwaysOnTop.Checked = true;
            this.miAlwaysOnTop.Index = 0;
            this.miAlwaysOnTop.Text = "Always on top";
            this.miAlwaysOnTop.Click += new System.EventHandler(this.miAlwaysOnTop_Click);
            // 
            // miOpacity
            // 
            this.miOpacity.Index = 1;
            this.miOpacity.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem3,
            this.menuItem4,
            this.menuItem5,
            this.menuItem6,
            this.menuItem7});
            this.miOpacity.Text = "Opacity 100%";
            // 
            // menuItem3
            // 
            this.menuItem3.Index = 0;
            this.menuItem3.Text = "20%";
            this.menuItem3.Click += new System.EventHandler(this.changeOpacity_Click);
            // 
            // menuItem4
            // 
            this.menuItem4.Index = 1;
            this.menuItem4.Text = "40%";
            this.menuItem4.Click += new System.EventHandler(this.changeOpacity_Click);
            // 
            // menuItem5
            // 
            this.menuItem5.Index = 2;
            this.menuItem5.Text = "60%";
            this.menuItem5.Click += new System.EventHandler(this.changeOpacity_Click);
            // 
            // menuItem6
            // 
            this.menuItem6.Index = 3;
            this.menuItem6.Text = "80%";
            this.menuItem6.Click += new System.EventHandler(this.changeOpacity_Click);
            // 
            // menuItem7
            // 
            this.menuItem7.Index = 4;
            this.menuItem7.Text = "100%";
            this.menuItem7.Click += new System.EventHandler(this.changeOpacity_Click);
            // 
            // miCopy
            // 
            this.miCopy.Index = 3;
            this.miCopy.Shortcut = System.Windows.Forms.Shortcut.CtrlC;
            this.miCopy.Text = "Copy";
            this.miCopy.Click += new System.EventHandler(this.miCopy_Click);
            // 
            // openXls
            // 
            this.openXls.DefaultExt = "*.xls";
            this.openXls.Filter = "Excel files (*.xls) |*.xls|All files|*.*";
            // 
            // MainForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(239, 249);
            this.Controls.Add(this.tvFields);
            this.Controls.Add(this.statusBar1);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Menu = this.MainMenu;
            this.Name = "MainForm";
            this.Text = "dCube Designer";
            this.TopMost = true;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion
	}
}

