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
    public partial class CompositeColumnsGroupcontrol : IEditGroupControl
	{

		private GridEXGroup mGroup;
		private GridEXTable mTable;
		private GridEXCustomGroup mCustomGroup;
		private GridEXCustomGroupRow mCustomGroupRow;
		private List<GridEXColumn> mCompositeColumns;

		public CompositeColumnsGroupcontrol(GridEXGroup group, GridEXTable table)
		{

			// This call is required by the Windows Form Designer.
			InitializeComponent();

			mCompositeColumns = new List<GridEXColumn>();

			this.CreateGridsTable();
			if (group.Table != null)
			{
				this.Table = group.Table;
			}
			else
			{
				this.Table = table;
			}
			// Add any initialization after the InitializeComponent() call.

			this.mGroup = group;
			if (mGroup.CustomGroup != null)
			{
				this.CustomGroup = mGroup.CustomGroup;
			}
			else
			{
				GridEXCustomGroup newCustomGroup = new GridEXCustomGroup();
				newCustomGroup.CustomGroupType = CustomGroupType.CompositeColumns;
				this.CustomGroup = newCustomGroup;
				mGroup.CustomGroup = newCustomGroup;
			}

		}

		private GridEXTable Table
		{
			get
			{
				return mTable;
			}
			set
			{
				if (mTable != value)
				{
					mTable = value;
					this.OnTableChanged();
				}
			}
		}

		private GridEXCustomGroup CustomGroup
		{
			get
			{
				return mCustomGroup;
			}
			set
			{
				mCustomGroup = value;
				this.OnCustomGroupChanged();
			}
		}


		public GridEXColumn[] CompositeColumns
		{
			get
			{
				return this.mCompositeColumns.ToArray();

			}
			set
			{
				if (value != null)
				{
					this.mCompositeColumns = new List<GridEXColumn>(value);
				}
				else
				{
					this.mCompositeColumns = new List<GridEXColumn>();
				}
				this.FillControls(true, true);
			}
		}

		private void OnTableChanged()
		{
			if (Table.ParentTable == null && Table.ChildTables.Count == 0)
			{
				this.grbTable.Visible = false;
			}
			else
			{
				this.grbTable.Visible = true;
			}

			this.FillControls(true, false);

		}

		private void OnCustomGroupChanged()
		{
			this.CompositeColumns = mCustomGroup.CompositeColumns;
		}

		private void CreateGridsTable()
		{
			this.grdColumnList.BoundMode = BoundMode.Unbound;
			this.grdColumnList.RootTable = new GridEXTable();

			GridEXColumn textColumn = new GridEXColumn();
			textColumn.Selectable = false;
			textColumn.DataMember = "Text";
			textColumn.Key = "Text";
			GridEXColumn valueColumn = new GridEXColumn();
			valueColumn.DataMember = "Value";
			valueColumn.Key = "Value";
			valueColumn.Visible = false;

			this.grdColumnList.RootTable.Columns.Add(textColumn);
			this.grdColumnList.RootTable.Columns.Add(valueColumn);

			this.grdCompositeColumns.BoundMode = BoundMode.Unbound;
			this.grdCompositeColumns.RootTable = new GridEXTable();

			textColumn = new GridEXColumn();
			textColumn.DataMember = "Text";
			textColumn.Key = "Text";
			textColumn.Selectable = false;
			valueColumn = new GridEXColumn();
			valueColumn.DataMember = "Value";
			valueColumn.Key = "Value";
			valueColumn.Visible = false;
			this.grdCompositeColumns.RootTable.Columns.Add(textColumn);
			this.grdCompositeColumns.RootTable.Columns.Add(valueColumn);

			this.grdColumnList.ColumnAutoResize = true;
			this.grdCompositeColumns.ColumnAutoResize = true;
		}

		private void FillControls(bool fillAvailable, bool fillSelected)
		{

			if (fillAvailable)
			{

				this.grdColumnList.ClearItems();
				if (this.mTable != null)
				{

					foreach (GridEXColumn col in mTable.Columns)
					{

						if (col.AllowGroup && ! col.MultipleValues && ! this.mCompositeColumns.Contains(col))
						{
							this.grdColumnList.AddItem(MainQD.GetColumnFriendlyName(col), col);
						}
					}
					if (this.grdColumnList.RecordCount > 0)
					{
						this.grdColumnList.MoveFirst();
					}
				}
			}
			if (fillSelected)
			{

				this.grdCompositeColumns.ClearItems();
				foreach (GridEXColumn col in this.CompositeColumns)
				{
					this.grdCompositeColumns.AddItem(MainQD.GetColumnFriendlyName(col), col);
				}
				if (this.grdCompositeColumns.RecordCount > 0)
				{
					this.grdCompositeColumns.MoveFirst();
				}

				this.CustomGroup.CompositeColumns = this.CompositeColumns;
				if (RefreshData != null)
					RefreshData(this, EventArgs.Empty);
			}
			this.EnableButtons();

		}


		private void EnableButtons()
		{
			this.btnAdd.Enabled = (this.grdColumnList.Row >= 0);
			this.btnMoveDown.Enabled = (this.grdCompositeColumns.Row < this.grdCompositeColumns.RecordCount - 1);
			this.btnMoveUp.Enabled = (this.grdCompositeColumns.Row > 0);
			this.btnRemove.Enabled = (this.grdCompositeColumns.Row >= 0);

		}

		private void DoAdd()
		{

			if (this.grdColumnList.GetRow() != null)
			{
				int index = this.grdColumnList.Row;
				this.mCompositeColumns.Add((GridEXColumn)this.grdColumnList.GetValue("Value"));
				this.FillControls(true, true);
				if (index < this.grdColumnList.RecordCount - 1)
				{
					this.grdColumnList.MoveTo(index);
				}
				else
				{
					this.grdColumnList.MoveLast();
				}

			}
		}
		private void DoRemove()
		{

			if (this.grdCompositeColumns.GetRow() != null)
			{
				int index = this.grdCompositeColumns.Row;
				this.mCompositeColumns.Remove((GridEXColumn)this.grdCompositeColumns.GetValue("Value"));
				this.FillControls(true, true);
				if (index < this.grdCompositeColumns.RecordCount - 1)
				{

					this.grdCompositeColumns.MoveTo(index);
				}
				else
				{
					this.grdCompositeColumns.MoveLast();
				}
			}
		}

		private void btnAdd_Click(object sender, System.EventArgs e)
		{

			DoAdd();
		}

		private void btnRemove_Click(object sender, System.EventArgs e)
		{
			DoRemove();

		}

		private void grdColumnList_RowDoubleClick(object sender, Janus.Windows.GridEX.RowActionEventArgs e)
		{
			DoAdd();
		}

		private void grdCompositeColumns_RowDoubleClick(object sender, Janus.Windows.GridEX.RowActionEventArgs e)
		{
			DoRemove();

		}

		private void btnMoveUp_Click(object sender, System.EventArgs e)
		{

			int index = this.grdCompositeColumns.Row;
			GridEXColumn col = this.CompositeColumns[index];

			this.mCompositeColumns.RemoveAt(index);

			this.mCompositeColumns.Insert(index - 1, col);
			this.FillControls(false, true);
			this.grdCompositeColumns.MoveTo(index - 1);

		}

		private void btnMoveDown_Click(object sender, System.EventArgs e)
		{

			int index = this.grdCompositeColumns.Row;
			GridEXColumn col = this.CompositeColumns[index];

			this.mCompositeColumns.RemoveAt(index);

			this.mCompositeColumns.Insert(index + 1, col);
			this.FillControls(false, true);
			this.grdCompositeColumns.MoveTo(index + 1);
		}

		private void grdColumnList_SelectionChanged(object sender, System.EventArgs e)
		{
			this.btnAdd.Enabled = (this.grdColumnList.GetRow() != null);
		}

		private void grdCompositeColumns_SelectionChanged(object sender, System.EventArgs e)
		{
			this.EnableButtons();
		}

        #region IEditGroupControl Members

        public event EventHandler RefreshData;

        #endregion
    }

} //end of root namespace