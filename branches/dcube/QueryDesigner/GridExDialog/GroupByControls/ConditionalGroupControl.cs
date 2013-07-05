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
    public partial class ConditionalGroupControl : IEditGroupControl
	{

		private GridEXGroup mGroup;
		private GridEXTable mTable;
		private GridEXCustomGroup mCustomGroup;
		private GridEXCustomGroupRow mCustomGroupRow;

		public ConditionalGroupControl(GridEXGroup group, GridEXTable table)
		{

			// This call is required by the Windows Form Designer.
			InitializeComponent();
			if (group.Table != null)
			{
				this.mTable = group.Table;
			}
			else
			{
				this.mTable = table;
			}
			// Add any initialization after the InitializeComponent() call.
			this.FillCustomGroupsCombo();

			this.Group = group;
		}

		protected override void OnLoad(System.EventArgs e)
		{
			base.OnLoad(e);
            //MainQD.MainForm.VisualStyleManager1.AddControl(this, true);
		}

		public GridEXGroup Group
		{
			get
			{
				return mGroup;
			}
			set
			{
				if (value != mGroup)
				{
					mGroup = value;
					this.OnGroupChanged();
				}
			}
		}

		public GridEXCustomGroup CustomGroup
		{
			get
			{
				return mCustomGroup;
			}
			set
			{
				if (value != mCustomGroup)
				{
					mCustomGroup = value;
					this.OnCustomGroupChanged();
				}
			}
		}

		public GridEXCustomGroupRow SelectedCustomGroupRow
		{
			get
			{
				return mCustomGroupRow;
			}
			set
			{
				if (value != mCustomGroupRow)
				{
					mCustomGroupRow = value;
					this.OnCustomGroupRowChanged();
				}
			}
		}

		private void FillCustomGroupsCombo()
		{
			foreach (GridEXCustomGroup customGroup in mTable.CustomGroups)
			{
				if (customGroup.CustomGroupType == CustomGroupType.ConditionalGroupRows)
				{
					this.cboSelectCustomGroup.Items.Add(customGroup.Key, customGroup);
				}
			}
		}


		private void FillGroupRowList()
		{

			grdGroupRows.RootTable = new GridEXTable();

			GridEXColumn column = new GridEXColumn();
			column.DataMember = "GroupCaption";
			column.Selectable = false;


			grdGroupRows.RootTable.Columns.Add(column);
			grdGroupRows.ColumnAutoResize = true;

			if (mCustomGroup != null)
			{
				this.grdGroupRows.SetDataBinding(mCustomGroup.GroupRows, "");
			}
			else
			{
				this.grdGroupRows.DataSource = null;
			}

		}

		private void OnCustomGroupRowChanged()
		{
			this.FilterEditor1.Table = mTable;
			bool controlsEnabled = (mCustomGroupRow != null);

			if (controlsEnabled)
			{
				this.FilterEditor1.FilterCondition = mCustomGroupRow.Condition;
				this.txtGroupRowCaption.Text = mCustomGroupRow.GroupCaption;
				this.chkShowWhenEmpty.Checked = mCustomGroupRow.ShowWhenEmpty;
			}
			else
			{
				this.FilterEditor1.FilterCondition = null;
				this.txtGroupRowCaption.Text = "";
			}

			this.FilterEditor1.Enabled = controlsEnabled;
			this.txtGroupRowCaption.Enabled = controlsEnabled;
			this.lblGroupRowCaption.Enabled = controlsEnabled;
			this.chkShowWhenEmpty.Enabled = controlsEnabled;

			this.btnRemoveGroupRow.Enabled = controlsEnabled;
			this.btnMoveDown.Enabled = controlsEnabled && (this.SelectedCustomGroupRow.Index < this.CustomGroup.GroupRows.Count - 1);
			this.btnMoveUp.Enabled = controlsEnabled && (this.SelectedCustomGroupRow.Index > 0);

		}

		private void OnCustomGroupChanged()
		{
			this.FillGroupRowList();
			this.cboSelectCustomGroup.SelectedValue = this.CustomGroup;

			this.Group.CustomGroup = this.CustomGroup;

			bool controlsEnabled = (mCustomGroup != null);

			if (mCustomGroup != null)
			{
				this.txtName.Text = mCustomGroup.Key;
			}
			else
			{
				this.txtName.Text = "";
			}

			this.lblName.Enabled = controlsEnabled;
			this.txtName.Enabled = controlsEnabled;
			this.grdGroupRows.Enabled = controlsEnabled;
			this.btnNewGroupRow.Enabled = controlsEnabled;
		}

		private void OnGroupChanged()
		{
			if (mGroup != null)
			{
				this.txtHeaderCaption.Text = mGroup.HeaderCaption;

				this.CustomGroup = mGroup.CustomGroup;
			}
		}

		private void CreateCustomGroup()
		{
			GridEXCustomGroup customGroup = new GridEXCustomGroup();
			customGroup.CustomGroupType = CustomGroupType.ConditionalGroupRows;
			customGroup.Key = MainQD.GetCustomGroupName(this.mTable.CustomGroups, "Custom Group ", 0);
			this.mTable.CustomGroups.Add(customGroup);

			this.cboSelectCustomGroup.Items.Add(customGroup.Key, customGroup);
			this.cboSelectCustomGroup.SelectedItem = this.cboSelectCustomGroup.Items[customGroup.Key];
		}

		private void btnNewCustomGroup_Click(object sender, System.EventArgs e)
		{
			this.cboSelectCustomGroup.SelectedItem = null;

			this.CreateCustomGroup();
		}

		private void cboSelectCustomGroup_SelectedItemChanged(object sender, System.EventArgs e)
		{
			if (cboSelectCustomGroup.SelectedItem != null)
			{
				this.CustomGroup = (GridEXCustomGroup)this.cboSelectCustomGroup.SelectedItem.DataRow;
			}
		}

		private void grdGroupRows_SelectionChanged(object sender, System.EventArgs e)
		{
			if (this.grdGroupRows.Row >= 0)
			{
                this.SelectedCustomGroupRow = (GridEXCustomGroupRow)this.grdGroupRows.GetRow().DataRow;
			}
			else
			{
				this.SelectedCustomGroupRow = null;
			}
		}

		private void btnNewGroupRow_Click(object sender, System.EventArgs e)
		{
			GridEXCustomGroupRow groupRow = new GridEXCustomGroupRow();

			this.CustomGroup.GroupRows.Add(groupRow);
			groupRow.GroupCaption = MainQD.GetCustomGroupRowName(this.CustomGroup.GroupRows, "Custom GroupRow ", 0);

			this.grdGroupRows.Refetch();
			this.grdGroupRows.MoveLast();

		}

		private void btnRemoveGroupRow_Click(object sender, System.EventArgs e)
		{

			if (this.SelectedCustomGroupRow != null)
			{
				this.grdGroupRows.AllowDelete = InheritableBoolean.True;
				this.grdGroupRows.Delete();
				this.grdGroupRows.AllowDelete = InheritableBoolean.False;
			}
		}

		private void btnMoveUp_Click(object sender, System.EventArgs e)
		{
			if (this.SelectedCustomGroupRow != null)
			{
				int index = this.SelectedCustomGroupRow.Index;
				GridEXCustomGroupRow groupRow = this.SelectedCustomGroupRow;

				this.CustomGroup.GroupRows.RemoveAt(index);

				this.CustomGroup.GroupRows.Insert(index - 1, groupRow);
				this.grdGroupRows.Refetch();
				this.grdGroupRows.MoveTo(index - 1);

			}
		}

		private void btnMoveDown_Click(object sender, System.EventArgs e)
		{
			if (this.SelectedCustomGroupRow != null)
			{
				int index = this.SelectedCustomGroupRow.Index;
				GridEXCustomGroupRow groupRow = this.SelectedCustomGroupRow;

				this.CustomGroup.GroupRows.RemoveAt(index);

				this.CustomGroup.GroupRows.Insert(index + 1, groupRow);
				this.grdGroupRows.Refetch();
				this.grdGroupRows.MoveTo(index + 1);
			}
		}

		private void chkShowWhenEmpty_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.SelectedCustomGroupRow != null)
			{
				this.SelectedCustomGroupRow.ShowWhenEmpty = this.chkShowWhenEmpty.Checked;
			}
		}

		private void txtGroupRowCaption_TextChanged(object sender, System.EventArgs e)
		{
			if (this.SelectedCustomGroupRow != null)
			{
				this.SelectedCustomGroupRow.GroupCaption = this.txtGroupRowCaption.Text;
				this.grdGroupRows.Refresh();
			}
		}

		private void txtHeaderCaption_TextChanged(object sender, System.EventArgs e)
		{
			if (mGroup != null)
			{
				this.Group.HeaderCaption = txtHeaderCaption.Text;
				if (RefreshData != null)
					RefreshData(this, EventArgs.Empty);
			}
		}

		private void txtName_TextChanged(object sender, System.EventArgs e)
		{
			if (mCustomGroup != null)
			{
				this.CustomGroup.Key = txtName.Text;
			}
		}

		private void FilterEditor1_FilterConditionChanged(object sender, System.EventArgs e)
		{
			if (this.SelectedCustomGroupRow != null)
			{
				this.SelectedCustomGroupRow.Condition = (GridEXFilterCondition)this.FilterEditor1.FilterCondition;
			}
		}


        #region IEditGroupControl Members

        public event EventHandler RefreshData;

        #endregion
    }

} //end of root namespace