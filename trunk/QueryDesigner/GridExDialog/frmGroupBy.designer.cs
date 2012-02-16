using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;

using Janus.Windows.GridEX;
using System.IO;

namespace QueryDesigner
{
	public partial class frmGroupBy : System.Windows.Forms.Form
	{

		//Form overrides dispose to clean up the component list.
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmGroupBy));
            this.UiCommandManager1 = new Janus.Windows.UI.CommandBars.UICommandManager(this.components);
            this.BottomRebar1 = new Janus.Windows.UI.CommandBars.UIRebar();
            this.UiCommandBar1 = new Janus.Windows.UI.CommandBars.UICommandBar();
            this.cmdTableList1 = new Janus.Windows.UI.CommandBars.UICommand("cmdTableList");
            this.cmdNew1 = new Janus.Windows.UI.CommandBars.UICommand("cmdNew");
            this.cmdRemove1 = new Janus.Windows.UI.CommandBars.UICommand("cmdRemove");
            this.cmdMoveUp1 = new Janus.Windows.UI.CommandBars.UICommand("cmdMoveUp");
            this.cmdMoveDown1 = new Janus.Windows.UI.CommandBars.UICommand("cmdMoveDown");
            this.cmdExpand1 = new Janus.Windows.UI.CommandBars.UICommand("cmdExpand");
            this.cmdHierarchicalGroupMode1 = new Janus.Windows.UI.CommandBars.UICommand("cmdHierarchicalGroupMode");
            this.cmdNew = new Janus.Windows.UI.CommandBars.UICommand("cmdNew");
            this.cmdNewSimpleGroup1 = new Janus.Windows.UI.CommandBars.UICommand("cmdNewSimpleGroup");
            this.cmdCompositeColumnsGroup1 = new Janus.Windows.UI.CommandBars.UICommand("cmdCompositeColumnsGroup");
            this.cmdNewCustomGroup1 = new Janus.Windows.UI.CommandBars.UICommand("cmdNewCustomGroup");
            this.cmdRemove = new Janus.Windows.UI.CommandBars.UICommand("cmdRemove");
            this.cmdNewCustomGroup = new Janus.Windows.UI.CommandBars.UICommand("cmdNewCustomGroup");
            this.cmdTableList = new Janus.Windows.UI.CommandBars.UICommand("cmdTableList");
            this.cmdNewSimpleGroup = new Janus.Windows.UI.CommandBars.UICommand("cmdNewSimpleGroup");
            this.cmdCompositeColumnsGroup = new Janus.Windows.UI.CommandBars.UICommand("cmdCompositeColumnsGroup");
            this.cmdMoveUp = new Janus.Windows.UI.CommandBars.UICommand("cmdMoveUp");
            this.cmdMoveDown = new Janus.Windows.UI.CommandBars.UICommand("cmdMoveDown");
            this.cmdExpand = new Janus.Windows.UI.CommandBars.UICommand("cmdExpand");
            this.cmdHierarchicalGroupMode = new Janus.Windows.UI.CommandBars.UICommand("cmdHierarchicalGroupMode");
            this.ImageList1 = new System.Windows.Forms.ImageList(this.components);
            this.LeftRebar1 = new Janus.Windows.UI.CommandBars.UIRebar();
            this.RightRebar1 = new Janus.Windows.UI.CommandBars.UIRebar();
            this.TopRebar1 = new Janus.Windows.UI.CommandBars.UIRebar();
            this.grdGroupList = new Janus.Windows.GridEX.GridEX();
            this.UiPanelManager1 = new Janus.Windows.UI.Dock.UIPanelManager(this.components);
            this.uiPanel0 = new Janus.Windows.UI.Dock.UIPanel();
            this.uiPanel0Container = new Janus.Windows.UI.Dock.UIPanelInnerContainer();
            this.uiPanel1 = new Janus.Windows.UI.Dock.UIPanel();
            this.uiPanel1Container = new Janus.Windows.UI.Dock.UIPanelInnerContainer();
            this.btnOK = new Janus.Windows.EditControls.UIButton();
            this.btnCancel = new Janus.Windows.EditControls.UIButton();
            this.officeFormAdorner1 = new Janus.Windows.Ribbon.OfficeFormAdorner(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.UiCommandManager1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.BottomRebar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.UiCommandBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LeftRebar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RightRebar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TopRebar1)).BeginInit();
            this.TopRebar1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdGroupList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.UiPanelManager1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.uiPanel0)).BeginInit();
            this.uiPanel0.SuspendLayout();
            this.uiPanel0Container.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.uiPanel1)).BeginInit();
            this.uiPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.officeFormAdorner1)).BeginInit();
            this.SuspendLayout();
            // 
            // UiCommandManager1
            // 
            this.UiCommandManager1.BottomRebar = this.BottomRebar1;
            this.UiCommandManager1.CommandBars.AddRange(new Janus.Windows.UI.CommandBars.UICommandBar[] {
            this.UiCommandBar1});
            this.UiCommandManager1.Commands.AddRange(new Janus.Windows.UI.CommandBars.UICommand[] {
            this.cmdNew,
            this.cmdRemove,
            this.cmdNewCustomGroup,
            this.cmdTableList,
            this.cmdNewSimpleGroup,
            this.cmdCompositeColumnsGroup,
            this.cmdMoveUp,
            this.cmdMoveDown,
            this.cmdExpand,
            this.cmdHierarchicalGroupMode});
            this.UiCommandManager1.ContainerControl = this;
            this.UiCommandManager1.Id = new System.Guid("48b2ef4b-55f7-47df-aa54-e13d23fe4723");
            this.UiCommandManager1.ImageList = this.ImageList1;
            this.UiCommandManager1.LeftRebar = this.LeftRebar1;
            this.UiCommandManager1.LockCommandBars = true;
            this.UiCommandManager1.RightRebar = this.RightRebar1;
            this.UiCommandManager1.ShowCustomizeButton = Janus.Windows.UI.InheritableBoolean.False;
            this.UiCommandManager1.Tag = null;
            this.UiCommandManager1.TopRebar = this.TopRebar1;
            this.UiCommandManager1.CommandClick += new Janus.Windows.UI.CommandBars.CommandEventHandler(this.UiCommandManager1_CommandClick);
            // 
            // BottomRebar1
            // 
            this.BottomRebar1.CommandManager = this.UiCommandManager1;
            this.BottomRebar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.BottomRebar1.Location = new System.Drawing.Point(0, 390);
            this.BottomRebar1.Name = "BottomRebar1";
            this.BottomRebar1.Size = new System.Drawing.Size(710, 0);
            // 
            // UiCommandBar1
            // 
            this.UiCommandBar1.CommandManager = this.UiCommandManager1;
            this.UiCommandBar1.Commands.AddRange(new Janus.Windows.UI.CommandBars.UICommand[] {
            this.cmdTableList1,
            this.cmdNew1,
            this.cmdRemove1,
            this.cmdMoveUp1,
            this.cmdMoveDown1,
            this.cmdExpand1,
            this.cmdHierarchicalGroupMode1});
            this.UiCommandBar1.FullRow = true;
            this.UiCommandBar1.Key = "CommandBar1";
            this.UiCommandBar1.Location = new System.Drawing.Point(0, 0);
            this.UiCommandBar1.Name = "UiCommandBar1";
            this.UiCommandBar1.RowIndex = 0;
            this.UiCommandBar1.Size = new System.Drawing.Size(857, 28);
            this.UiCommandBar1.Text = "CommandBar1";
            // 
            // cmdTableList1
            // 
            this.cmdTableList1.Key = "cmdTableList";
            this.cmdTableList1.Name = "cmdTableList1";
            // 
            // cmdNew1
            // 
            this.cmdNew1.Key = "cmdNew";
            this.cmdNew1.Name = "cmdNew1";
            // 
            // cmdRemove1
            // 
            this.cmdRemove1.Key = "cmdRemove";
            this.cmdRemove1.Name = "cmdRemove1";
            // 
            // cmdMoveUp1
            // 
            this.cmdMoveUp1.Key = "cmdMoveUp";
            this.cmdMoveUp1.Name = "cmdMoveUp1";
            // 
            // cmdMoveDown1
            // 
            this.cmdMoveDown1.Key = "cmdMoveDown";
            this.cmdMoveDown1.Name = "cmdMoveDown1";
            // 
            // cmdExpand1
            // 
            this.cmdExpand1.Key = "cmdExpand";
            this.cmdExpand1.Name = "cmdExpand1";
            // 
            // cmdHierarchicalGroupMode1
            // 
            this.cmdHierarchicalGroupMode1.Key = "cmdHierarchicalGroupMode";
            this.cmdHierarchicalGroupMode1.Name = "cmdHierarchicalGroupMode1";
            // 
            // cmdNew
            // 
            this.cmdNew.Commands.AddRange(new Janus.Windows.UI.CommandBars.UICommand[] {
            this.cmdNewSimpleGroup1,
            this.cmdCompositeColumnsGroup1,
            this.cmdNewCustomGroup1});
            this.cmdNew.ImageIndex = 9;
            this.cmdNew.Key = "cmdNew";
            this.cmdNew.Name = "cmdNew";
            this.cmdNew.Text = "New Group";
            // 
            // cmdNewSimpleGroup1
            // 
            this.cmdNewSimpleGroup1.Key = "cmdNewSimpleGroup";
            this.cmdNewSimpleGroup1.Name = "cmdNewSimpleGroup1";
            // 
            // cmdCompositeColumnsGroup1
            // 
            this.cmdCompositeColumnsGroup1.Key = "cmdCompositeColumnsGroup";
            this.cmdCompositeColumnsGroup1.Name = "cmdCompositeColumnsGroup1";
            // 
            // cmdNewCustomGroup1
            // 
            this.cmdNewCustomGroup1.Key = "cmdNewCustomGroup";
            this.cmdNewCustomGroup1.Name = "cmdNewCustomGroup1";
            // 
            // cmdRemove
            // 
            this.cmdRemove.ImageIndex = 12;
            this.cmdRemove.Key = "cmdRemove";
            this.cmdRemove.Name = "cmdRemove";
            this.cmdRemove.Text = "Remove Group";
            // 
            // cmdNewCustomGroup
            // 
            this.cmdNewCustomGroup.ImageIndex = 1;
            this.cmdNewCustomGroup.Key = "cmdNewCustomGroup";
            this.cmdNewCustomGroup.Name = "cmdNewCustomGroup";
            this.cmdNewCustomGroup.Text = "New Conditional Group";
            // 
            // cmdTableList
            // 
            this.cmdTableList.CommandType = Janus.Windows.UI.CommandBars.CommandType.ComboBoxCommand;
            this.cmdTableList.IsEditableControl = Janus.Windows.UI.InheritableBoolean.True;
            this.cmdTableList.Key = "cmdTableList";
            this.cmdTableList.Name = "cmdTableList";
            this.cmdTableList.ShowTextInContainers = Janus.Windows.UI.InheritableBoolean.True;
            this.cmdTableList.Text = "Groups in:";
            // 
            // cmdNewSimpleGroup
            // 
            this.cmdNewSimpleGroup.ImageIndex = 3;
            this.cmdNewSimpleGroup.Key = "cmdNewSimpleGroup";
            this.cmdNewSimpleGroup.Name = "cmdNewSimpleGroup";
            this.cmdNewSimpleGroup.Text = "New Group By Field";
            // 
            // cmdCompositeColumnsGroup
            // 
            this.cmdCompositeColumnsGroup.ImageIndex = 7;
            this.cmdCompositeColumnsGroup.Key = "cmdCompositeColumnsGroup";
            this.cmdCompositeColumnsGroup.Name = "cmdCompositeColumnsGroup";
            this.cmdCompositeColumnsGroup.Text = "New Group By Multiple Fields";
            // 
            // cmdMoveUp
            // 
            this.cmdMoveUp.ImageIndex = 6;
            this.cmdMoveUp.Key = "cmdMoveUp";
            this.cmdMoveUp.Name = "cmdMoveUp";
            this.cmdMoveUp.Text = "Move Up";
            // 
            // cmdMoveDown
            // 
            this.cmdMoveDown.ImageIndex = 5;
            this.cmdMoveDown.Key = "cmdMoveDown";
            this.cmdMoveDown.Name = "cmdMoveDown";
            this.cmdMoveDown.Text = "Move Down";
            // 
            // cmdExpand
            // 
            this.cmdExpand.Checked = Janus.Windows.UI.InheritableBoolean.True;
            this.cmdExpand.CommandType = Janus.Windows.UI.CommandBars.CommandType.ToggleButton;
            this.cmdExpand.ImageIndex = 2;
            this.cmdExpand.Key = "cmdExpand";
            this.cmdExpand.Name = "cmdExpand";
            this.cmdExpand.Text = "All Expanded";
            // 
            // cmdHierarchicalGroupMode
            // 
            this.cmdHierarchicalGroupMode.Alignment = Janus.Windows.UI.CommandBars.CommandAlignment.Far;
            this.cmdHierarchicalGroupMode.CommandType = Janus.Windows.UI.CommandBars.CommandType.ComboBoxCommand;
            this.cmdHierarchicalGroupMode.IsEditableControl = Janus.Windows.UI.InheritableBoolean.True;
            this.cmdHierarchicalGroupMode.Key = "cmdHierarchicalGroupMode";
            this.cmdHierarchicalGroupMode.Name = "cmdHierarchicalGroupMode";
            this.cmdHierarchicalGroupMode.ShowTextInContainers = Janus.Windows.UI.InheritableBoolean.True;
            this.cmdHierarchicalGroupMode.Text = "Group Mode";
            // 
            // ImageList1
            // 
            this.ImageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("ImageList1.ImageStream")));
            this.ImageList1.TransparentColor = System.Drawing.Color.Magenta;
            this.ImageList1.Images.SetKeyName(0, "");
            this.ImageList1.Images.SetKeyName(1, "");
            this.ImageList1.Images.SetKeyName(2, "");
            this.ImageList1.Images.SetKeyName(3, "");
            this.ImageList1.Images.SetKeyName(4, "");
            this.ImageList1.Images.SetKeyName(5, "");
            this.ImageList1.Images.SetKeyName(6, "");
            this.ImageList1.Images.SetKeyName(7, "");
            this.ImageList1.Images.SetKeyName(8, "");
            this.ImageList1.Images.SetKeyName(9, "");
            this.ImageList1.Images.SetKeyName(10, "");
            this.ImageList1.Images.SetKeyName(11, "");
            this.ImageList1.Images.SetKeyName(12, "");
            // 
            // LeftRebar1
            // 
            this.LeftRebar1.CommandManager = this.UiCommandManager1;
            this.LeftRebar1.Dock = System.Windows.Forms.DockStyle.Left;
            this.LeftRebar1.Location = new System.Drawing.Point(0, 0);
            this.LeftRebar1.Name = "LeftRebar1";
            this.LeftRebar1.Size = new System.Drawing.Size(0, 390);
            // 
            // RightRebar1
            // 
            this.RightRebar1.CommandManager = this.UiCommandManager1;
            this.RightRebar1.Dock = System.Windows.Forms.DockStyle.Right;
            this.RightRebar1.Location = new System.Drawing.Point(710, 0);
            this.RightRebar1.Name = "RightRebar1";
            this.RightRebar1.Size = new System.Drawing.Size(0, 390);
            // 
            // TopRebar1
            // 
            this.TopRebar1.CommandBars.AddRange(new Janus.Windows.UI.CommandBars.UICommandBar[] {
            this.UiCommandBar1});
            this.TopRebar1.CommandManager = this.UiCommandManager1;
            this.TopRebar1.Controls.Add(this.UiCommandBar1);
            this.TopRebar1.Dock = System.Windows.Forms.DockStyle.Top;
            this.TopRebar1.Location = new System.Drawing.Point(0, 0);
            this.TopRebar1.Name = "TopRebar1";
            this.TopRebar1.Size = new System.Drawing.Size(857, 28);
            // 
            // grdGroupList
            // 
            this.grdGroupList.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.False;
            this.grdGroupList.BorderStyle = Janus.Windows.GridEX.BorderStyle.None;
            this.grdGroupList.ColumnHeaders = Janus.Windows.GridEX.InheritableBoolean.False;
            this.grdGroupList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grdGroupList.GridLines = Janus.Windows.GridEX.GridLines.None;
            this.grdGroupList.GroupByBoxVisible = false;
            this.grdGroupList.HideSelection = Janus.Windows.GridEX.HideSelection.Highlight;
            this.grdGroupList.ImageList = this.ImageList1;
            this.grdGroupList.Location = new System.Drawing.Point(0, 0);
            this.grdGroupList.Name = "grdGroupList";
            this.grdGroupList.Size = new System.Drawing.Size(161, 315);
            this.grdGroupList.TabIndex = 0;
            this.grdGroupList.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003;
            this.grdGroupList.SelectionChanged += new System.EventHandler(this.grdGroupList_SelectionChanged);
            this.grdGroupList.CurrentCellChanging += new Janus.Windows.GridEX.CurrentCellChangingEventHandler(this.grdGroupList_CurrentCellChanging);
            this.grdGroupList.FormattingRow += new Janus.Windows.GridEX.RowLoadEventHandler(this.grdGroupList_FormattingRow);
            // 
            // UiPanelManager1
            // 
            this.UiPanelManager1.ContainerControl = this;
            this.UiPanelManager1.DefaultPanelSettings.CaptionVisible = false;
            this.UiPanelManager1.PanelPadding.Bottom = 45;
            this.UiPanelManager1.PanelPadding.Left = 4;
            this.UiPanelManager1.PanelPadding.Right = 4;
            this.UiPanelManager1.PanelPadding.Top = 4;
            this.UiPanelManager1.Tag = null;
            this.uiPanel0.Id = new System.Guid("3d65b9ea-5b85-478d-b855-d837f98a5649");
            this.UiPanelManager1.Panels.Add(this.uiPanel0);
            this.uiPanel1.Id = new System.Guid("809ccdb1-da44-4be9-8362-3863df05aec2");
            this.UiPanelManager1.Panels.Add(this.uiPanel1);
            // 
            // Design Time Panel Info:
            // 
            this.UiPanelManager1.BeginPanelInfo();
            this.UiPanelManager1.AddDockPanelInfo(new System.Guid("3d65b9ea-5b85-478d-b855-d837f98a5649"), Janus.Windows.UI.Dock.PanelDockStyle.Left, new System.Drawing.Size(167, 317), true);
            this.UiPanelManager1.AddDockPanelInfo(new System.Guid("809ccdb1-da44-4be9-8362-3863df05aec2"), Janus.Windows.UI.Dock.PanelDockStyle.Fill, new System.Drawing.Size(682, 317), true);
            this.UiPanelManager1.AddFloatingPanelInfo(new System.Guid("3d65b9ea-5b85-478d-b855-d837f98a5649"), new System.Drawing.Point(-1, -1), new System.Drawing.Size(-1, -1), false);
            this.UiPanelManager1.AddFloatingPanelInfo(new System.Guid("809ccdb1-da44-4be9-8362-3863df05aec2"), new System.Drawing.Point(-1, -1), new System.Drawing.Size(-1, -1), false);
            this.UiPanelManager1.EndPanelInfo();
            // 
            // uiPanel0
            // 
            this.uiPanel0.InnerContainer = this.uiPanel0Container;
            this.uiPanel0.Location = new System.Drawing.Point(4, 32);
            this.uiPanel0.Name = "uiPanel0";
            this.uiPanel0.Size = new System.Drawing.Size(167, 317);
            this.uiPanel0.TabIndex = 4;
            this.uiPanel0.Text = "Panel 0";
            // 
            // uiPanel0Container
            // 
            this.uiPanel0Container.Controls.Add(this.grdGroupList);
            this.uiPanel0Container.Location = new System.Drawing.Point(1, 1);
            this.uiPanel0Container.Name = "uiPanel0Container";
            this.uiPanel0Container.Size = new System.Drawing.Size(161, 315);
            this.uiPanel0Container.TabIndex = 0;
            this.uiPanel0Container.TabStop = false;
            // 
            // uiPanel1
            // 
            this.uiPanel1.InnerAreaStyle = Janus.Windows.UI.Dock.PanelInnerAreaStyle.ContainerPanel;
            this.uiPanel1.InnerContainer = this.uiPanel1Container;
            this.uiPanel1.Location = new System.Drawing.Point(171, 32);
            this.uiPanel1.Name = "uiPanel1";
            this.uiPanel1.Size = new System.Drawing.Size(682, 317);
            this.uiPanel1.TabIndex = 4;
            this.uiPanel1.Text = "Panel 1";
            // 
            // uiPanel1Container
            // 
            this.uiPanel1Container.Location = new System.Drawing.Point(1, 1);
            this.uiPanel1Container.Name = "uiPanel1Container";
            this.uiPanel1Container.Size = new System.Drawing.Size(680, 315);
            this.uiPanel1Container.TabIndex = 0;
            this.uiPanel1Container.TabStop = false;
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(682, 357);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 5;
            this.btnOK.Text = "OK";
            this.btnOK.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(763, 357);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // officeFormAdorner1
            // 
            this.officeFormAdorner1.DocumentName = "Group By...";
            this.officeFormAdorner1.Form = this;
            this.officeFormAdorner1.Office2007CustomColor = System.Drawing.Color.Empty;
            // 
            // frmGroupBy
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(857, 394);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.uiPanel1);
            this.Controls.Add(this.uiPanel0);
            this.Controls.Add(this.TopRebar1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmGroupBy";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Group By...";
            ((System.ComponentModel.ISupportInitialize)(this.UiCommandManager1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.BottomRebar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.UiCommandBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LeftRebar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RightRebar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TopRebar1)).EndInit();
            this.TopRebar1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.grdGroupList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.UiPanelManager1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.uiPanel0)).EndInit();
            this.uiPanel0.ResumeLayout(false);
            this.uiPanel0Container.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.uiPanel1)).EndInit();
            this.uiPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.officeFormAdorner1)).EndInit();
            this.ResumeLayout(false);

		}
		internal Janus.Windows.UI.CommandBars.UICommandManager UiCommandManager1;
		internal Janus.Windows.UI.CommandBars.UIRebar BottomRebar1;
		internal Janus.Windows.GridEX.GridEX grdGroupList;
		internal Janus.Windows.UI.CommandBars.UIRebar LeftRebar1;
		internal Janus.Windows.UI.CommandBars.UIRebar RightRebar1;
		internal Janus.Windows.UI.CommandBars.UIRebar TopRebar1;
		internal Janus.Windows.UI.Dock.UIPanel uiPanel1;
		internal Janus.Windows.UI.Dock.UIPanelInnerContainer uiPanel1Container;
		internal Janus.Windows.UI.Dock.UIPanel uiPanel0;
		internal Janus.Windows.UI.Dock.UIPanelInnerContainer uiPanel0Container;
		internal Janus.Windows.UI.Dock.UIPanelManager UiPanelManager1;
		internal Janus.Windows.UI.CommandBars.UICommandBar UiCommandBar1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdNew1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdRemove1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdNew;
		internal Janus.Windows.UI.CommandBars.UICommand cmdRemove;
		internal Janus.Windows.UI.CommandBars.UICommand cmdNewCustomGroup1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdNewCustomGroup;
		internal Janus.Windows.EditControls.UIButton btnOK;
		internal Janus.Windows.EditControls.UIButton btnCancel;
		internal Janus.Windows.UI.CommandBars.UICommand cmdTableList1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdTableList;
		internal Janus.Windows.UI.CommandBars.UICommand cmdNewSimpleGroup1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdNewSimpleGroup;
		internal Janus.Windows.UI.CommandBars.UICommand cmdCompositeColumnsGroup;
		internal Janus.Windows.UI.CommandBars.UICommand cmdCompositeColumnsGroup1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdMoveUp1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdMoveDown1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdMoveUp;
		internal Janus.Windows.UI.CommandBars.UICommand cmdMoveDown;
		internal Janus.Windows.UI.CommandBars.UICommand cmdExpand1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdExpand;
		internal Janus.Windows.UI.CommandBars.UICommand cmdHierarchicalGroupMode1;
		internal Janus.Windows.UI.CommandBars.UICommand cmdHierarchicalGroupMode;
		internal System.Windows.Forms.ImageList ImageList1;
        private Janus.Windows.Ribbon.OfficeFormAdorner officeFormAdorner1;
	}

} //end of root namespace