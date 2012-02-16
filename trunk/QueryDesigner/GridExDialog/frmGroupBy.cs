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
	public partial class frmGroupBy
	{

		private System.Collections.ArrayList groupCollection;
		private GridEX bufferGrid;
		private UserControl mActiveControl;
		private GridEXGroup mSelectedGroup;
		private GridEXTable mTable;

		protected override void OnLoad(System.EventArgs e)
		{
			base.OnLoad(e);
            //MainQD.MainForm.VisualStyleManager1.AddControl(this, true);
		}

		public DialogResult ShowDialog(GridEX grid, Form parent)
		{

			MemoryStream layoutStream = new MemoryStream();

			grid.SaveLayoutFile(layoutStream);
			layoutStream.Flush();
			layoutStream.Position = 0;

			this.bufferGrid = new GridEX();
			this.bufferGrid.LoadLayoutFile(layoutStream);

			//Set the DataBinding of the Grid and DropDowns in order to be able to retrieve
			//the ValueList of the columns used in the FilterEditor
			this.bufferGrid.BindingContext = this.BindingContext;
			this.bufferGrid.SetDataBinding(grid.DataSource, grid.DataMember);
			this.bufferGrid.ImageList = grid.ImageList;

            for (int i = 0; i < grid.DropDowns.Count; i++)
			{
				GridEXDropDown ddMain = grid.DropDowns[i];
				GridEXDropDown ddBuffer = bufferGrid.DropDowns[i];
				ddBuffer.SetDataBinding(ddMain.DataSource, ddMain.DataMember);

			}

			layoutStream.Dispose();

			this.Table = bufferGrid.RootTable;
			this.FillTablesCombo();
			this.FillHierarchicalGroupModeCombo();

			if (bufferGrid.GroupMode == GroupMode.Collapsed)
			{
				this.UiCommandManager1.Commands["cmdExpand"].IsChecked = false;
				this.UiCommandManager1.Commands["cmdExpand"].Text = "All collapsed";
			}

			this.ShowDialog(parent);


			if (this.DialogResult == System.Windows.Forms.DialogResult.OK)
			{

                if (!this.CreateGroups())
                {
                    return this.DialogResult;
                }

				layoutStream = new MemoryStream();
				this.bufferGrid.SaveLayoutFile(layoutStream);

				layoutStream.Flush();
				layoutStream.Position = 0;

				grid.LoadLayoutFile(layoutStream);

                for (int i = 0; i < grid.DropDowns.Count; i++)
				{
					GridEXDropDown ddMain = grid.DropDowns[i];
					GridEXDropDown ddBuffer = bufferGrid.DropDowns[i];
					ddMain.SetDataBinding(ddBuffer.DataSource, ddBuffer.DataMember);
				}
				grid.Refetch();
			}
			return this.DialogResult;
		}

		public GridEXTable Table
		{
			get
			{
				return mTable;
			}
			set
			{
				if (mTable != value)
				{
					this.CreateGroups();
					mTable = value;
					this.OnTableChanged();

				}
			}
		}

		public GridEXGroup SelectedGroup
		{
			get
			{
				return mSelectedGroup;
			}
			set
			{
				if (mSelectedGroup != value)
				{
					mSelectedGroup = value;
					this.OnSelectedGroupChanged(mSelectedGroup);
				}
			}
		}

		public UserControl ActiveGroupControl
		{
			get
			{
				return mActiveControl;
			}
			set
			{
				if (mActiveControl != null)
				{
                    ((IEditGroupControl)mActiveControl).RefreshData -= new System.EventHandler(this.ActiveControl_RefreshData);
					mActiveControl.Dispose();
				}
				mActiveControl = value;
				if (mActiveControl != null)
				{
                    ((IEditGroupControl)mActiveControl).RefreshData += new System.EventHandler(this.ActiveControl_RefreshData);

					this.uiPanel1.InnerContainer.Controls.Add(mActiveControl);
				}
			}
		}

		private void OnTableChanged()
		{
			this.FillGroupList();
			if (mTable.HierarchicalMode != HierarchicalMode.UseChildTables)
			{
				this.UiCommandManager1.Commands["cmdHierarchicalGroupMode"].IsVisible = true;
				Janus.Windows.EditControls.UIComboBox combo = this.UiCommandManager1.Commands["cmdHierarchicalGroupMode"].GetUIComboBox();
				combo.SelectedValue = this.Table.SelfReferencingSettings.HierarchicalGroupMode;
			}
			else
			{
				this.UiCommandManager1.Commands["cmdHierarchicalGroupMode"].IsVisible = false;
			}

		}

		private void FillGroupList()
		{

			groupCollection = new ArrayList();

			foreach (GridEXGroup group in Table.Groups)
			{
				groupCollection.Add(group);
			}
			this.grdGroupList.RootTable = new GridEXTable();

			GridEXColumn column = new GridEXColumn();
			column.DataMember = "HeaderCaption";
			column.Key = "HeaderCaption";
			column.ColumnType = ColumnType.ImageAndText;
			column.Selectable = false;

			grdGroupList.RootTable.Columns.Add(column);
			grdGroupList.ColumnAutoResize = true;
			this.grdGroupList.SetDataBinding(groupCollection, "");

			this.EnableControls();
		}

		private void FillHierarchicalGroupModeCombo()
		{
			Janus.Windows.EditControls.UIComboBox combo = this.UiCommandManager1.Commands["cmdHierarchicalGroupMode"].GetUIComboBox();

			combo.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
			combo.Items.Add("All Rows", HierarchicalGroupMode.AllRows);
			combo.Items.Add("Parent Rows", HierarchicalGroupMode.ParentRows);

			combo.SelectedValue = this.Table.SelfReferencingSettings.HierarchicalGroupMode;

			combo.SelectedItemChanged += new System.EventHandler(this.cboHierarchicalGroupMode_SelectedItemChanged);

		}

		private void FillTablesCombo()
		{
			if (this.bufferGrid.RootTable.ChildTables.Count > 0)
			{
				Janus.Windows.EditControls.UIComboBox cboTables = this.UiCommandBar1.Commands["cmdTableList"].GetUIComboBox();
				cboTables.ComboStyle = Janus.Windows.EditControls.ComboStyle.DropDownList;
				cboTables.ImageList = this.bufferGrid.ImageList;
				this.AddChildTables(cboTables, this.bufferGrid.RootTable, 0);

				cboTables.SelectedValue = this.bufferGrid.RootTable;

				cboTables.SelectedItemChanged += new System.EventHandler(cboTables_SelectedItemChanged);
			}
			else
			{
				this.UiCommandManager1.Commands["cmdTableList"].IsVisible = false;
			}

		}

		private void AddChildTables(Janus.Windows.EditControls.UIComboBox cboTables, GridEXTable table, int indent)
		{
			Janus.Windows.EditControls.UIComboBoxItem item = new Janus.Windows.EditControls.UIComboBoxItem(table.Caption, table);
			item.ImageIndex = table.ImageIndex;
			item.IndentLevel = indent;

			cboTables.Items.Add(item);
			foreach (GridEXTable childTable in table.ChildTables)
			{
				this.AddChildTables(cboTables, childTable, indent + 1);
			}

		}

		private bool CreateGroups()
		{
            try
            {
                if (groupCollection != null)
                {
                    this.Table.Groups.Clear();

                    if (groupCollection.Count > 0)
                    {
                        GridEXGroup[] groups = new GridEXGroup[groupCollection.Count];
                        this.groupCollection.CopyTo(groups);
                        this.Table.Groups.AddRange(groups);
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, MainQD.MessageCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
		}

		private void EnableControls()
		{
			this.UiCommandManager1.Commands["cmdMoveUp"].IsEnabled = this.grdGroupList.Row > 0;
			this.UiCommandManager1.Commands["cmdMoveDown"].IsEnabled = this.grdGroupList.Row < this.grdGroupList.RecordCount - 1;
			this.UiCommandManager1.Commands["cmdRemove"].IsEnabled = this.SelectedGroup != null;
		}
        public static string MessageCaption = "Universal Query Designer";
		private bool ValidateGroup(GridEXGroup group)
		{
			if (group == null)
			{
				return true;
			}

			if (group.CustomGroup != null)
			{
				return ValidateCustomGroup(group.CustomGroup);
			}
			if (group.Column != null)
			{
				return true;
			}
			else
			{
				MessageBox.Show("Select a field for the group or remove the group", MainQD.MessageCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}

			return false;
		}

		private bool ValidateCustomGroup(GridEXCustomGroup customGroup)
		{

			if (customGroup.CustomGroupType == CustomGroupType.CompositeColumns)
			{
				if (customGroup.CompositeColumns == null || customGroup.CompositeColumns.Length == 0)
				{

					MessageBox.Show("CompositeColumns are not defined in the group", MainQD.MessageCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return false;
				}
				else
				{
					foreach (GridEXColumn col in customGroup.CompositeColumns)
					{
						if (! this.Table.CanGroupBy(col))
						{
							MessageBox.Show("Column '" + col.DataMember + "' can not be used in a group of this table.", MainQD.MessageCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
							return false;
						}
					}
				}
			}
			return true;
		}

		private void MoveUp()
		{
			if (this.grdGroupList.Row > 0)
			{
				int index = this.grdGroupList.Row;
				GridEXGroup group = (GridEXGroup)this.groupCollection[index];

				this.groupCollection.RemoveAt(index);

				this.groupCollection.Insert(index - 1, group);
				this.grdGroupList.Refetch();
				this.grdGroupList.MoveTo(index - 1);
			}
		}

		private void MoveDown()
		{
			if (this.grdGroupList.Row < this.groupCollection.Count - 1)
			{
				int index = this.grdGroupList.Row;
				GridEXGroup group = (GridEXGroup)this.groupCollection[index];

				this.groupCollection.RemoveAt(index);

				this.groupCollection.Insert(index + 1, group);
				this.grdGroupList.Refetch();
				this.grdGroupList.MoveTo(index + 1);
			}
		}

		private void ActiveControl_RefreshData(object sender, EventArgs e)
		{
			this.grdGroupList.Refresh();
		}

		private void cboTables_SelectedItemChanged(object sender, System.EventArgs e)
		{
			Janus.Windows.EditControls.UIComboBox combo = (Janus.Windows.EditControls.UIComboBox)sender;
			this.Table = (GridEXTable)combo.SelectedItem.Value;
		}

		private void cboHierarchicalGroupMode_SelectedItemChanged(object sender, System.EventArgs e)
		{
			if (this.Table != null)
			{
				Janus.Windows.EditControls.UIComboBox combo = (Janus.Windows.EditControls.UIComboBox)sender;
				this.Table.SelfReferencingSettings.HierarchicalGroupMode = (HierarchicalGroupMode)combo.SelectedItem.Value;
			}
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void grdGroupList_CurrentCellChanging(object sender, Janus.Windows.GridEX.CurrentCellChangingEventArgs e)
		{
			if (! ValidateGroup(this.SelectedGroup))
			{
				e.Cancel = true;
			}
		}

		private void grdGroupList_FormattingRow(object sender, Janus.Windows.GridEX.RowLoadEventArgs e)
		{
			if (e.Row.Cells.Count > 0)
			{
				if (e.Row.Cells["HeaderCaption"].Text.Length == 0)
				{
					e.Row.Cells["HeaderCaption"].Text = "New group";
				}
				GridEXGroup group = (GridEXGroup)e.Row.DataRow;
				if (group.CustomGroup != null)
				{
					if (group.CustomGroup.CustomGroupType == CustomGroupType.CompositeColumns)
					{
						e.Row.Cells["HeaderCaption"].ImageIndex = 7;
					}
					else
					{
						e.Row.Cells["HeaderCaption"].ImageIndex = 1;

					}
				}
				else
				{
					if (group.Column != null)
					{
						e.Row.Cells["HeaderCaption"].ImageIndex = 3;
					}
					else
					{
						e.Row.Cells["HeaderCaption"].ImageIndex = 4;
					}

				}

			}

		}

		private void grdGroupList_SelectionChanged(object sender, System.EventArgs e)
		{
			if (this.grdGroupList.Row >= 0)
			{
				this.SelectedGroup = (GridEXGroup)(this.grdGroupList.GetRow().DataRow);
			}
			else
			{
				this.SelectedGroup = null;
			}
		}

		private void UiCommandManager1_CommandClick(object sender, Janus.Windows.UI.CommandBars.CommandEventArgs e)
		{
			switch (e.Command.Key)
			{
				case "cmdNewCustomGroup":
					if (this.ValidateGroup(this.SelectedGroup))
					{
						this.CreateFilterGroupControl(null);
					}
					break;
				case "cmdNewSimpleGroup":
					if (this.ValidateGroup(this.SelectedGroup))
					{
						this.CreateColumnGroupControl(null);
					}
					break;
				case "cmdCompositeColumnsGroup":
					if (this.ValidateGroup(this.SelectedGroup))
					{
						this.CreateCompositeColumnsGroupControl(null);
					}
					break;
				case "cmdRemove":
					this.grdGroupList.AllowDelete = InheritableBoolean.True;
					this.grdGroupList.Delete();
					this.grdGroupList.AllowDelete = InheritableBoolean.False;
					break;
				case "cmdMoveUp":
					if (this.ValidateGroup(this.SelectedGroup))
					{
						this.MoveUp();
					}
					break;
				case "cmdMoveDown":
					if (this.ValidateGroup(this.SelectedGroup))
					{
						this.MoveDown();
					}
					break;
				case "cmdExpand":
					if (e.Command.IsChecked)
					{
						this.bufferGrid.GroupMode = GroupMode.Expanded;
						e.Command.Text = "All Expanded";
					}
					else
					{
						this.bufferGrid.GroupMode = GroupMode.Collapsed;
						e.Command.Text = "All Collapsed";
					}
					break;
			}
		}

		private void OnSelectedGroupChanged(GridEXGroup group)
		{
			if (group != null)
			{
				if (group.CustomGroup != null)
				{
					if (group.CustomGroup.CustomGroupType == CustomGroupType.ConditionalGroupRows)
					{
						this.CreateFilterGroupControl(group);
					}
					else
					{
						this.CreateCompositeColumnsGroupControl(group);
					}
				}
				else
				{
					this.CreateColumnGroupControl(group);
				}
			}
			else
			{
				this.ActiveGroupControl = null;
			}

			this.EnableControls();
		}

		private void CreateColumnGroupControl(GridEXGroup group)
		{
			if (group == null)
			{
				group = this.AddNewGroup();
			}
			ColumnGroupControl columnBasedControl = new ColumnGroupControl(group, this.Table);
			columnBasedControl.Dock = DockStyle.Fill;

			this.ActiveGroupControl = columnBasedControl;
		}

		private void CreateFilterGroupControl(GridEXGroup group)
		{
			if (group == null)
			{
				group = this.AddNewGroup();
			}
			ConditionalGroupControl filterBasedControl = new ConditionalGroupControl(group, this.Table);
			filterBasedControl.Dock = DockStyle.Fill;

			this.ActiveGroupControl = filterBasedControl;
		}

		private void CreateCompositeColumnsGroupControl(GridEXGroup group)
		{
			if (group == null)
			{
				group = this.AddNewGroup();
			}

			CompositeColumnsGroupcontrol compositeColumnsControl = new CompositeColumnsGroupcontrol(group, this.Table);
			compositeColumnsControl.Dock = DockStyle.Fill;

			this.ActiveGroupControl = compositeColumnsControl;

		}

		private GridEXGroup AddNewGroup()
		{
			GridEXGroup group = new GridEXGroup();
			this.groupCollection.Add(group);
			this.grdGroupList.Refetch();
			this.grdGroupList.MoveLast();
			return group;
		}

		public frmGroupBy()
		{

			// This call is required by the Windows Form Designer.
			InitializeComponent();

			// Add any initialization after the InitializeComponent() call.

		}
	}
} //end of root namespace