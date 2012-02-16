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
    public partial class ColumnGroupControl : IEditGroupControl
    {

        private GridEXGroup mGroup;
        private GridEXTable mTable;
        private GridEXColumn mColumn;

        public ColumnGroupControl(GridEXGroup group, GridEXTable table)
        {

            // This call is required by the Windows Form Designer.
            InitializeComponent();

            this.mGroup = group;
            this.FillGroupIntervalCombo();
            if (group.Column != null)
            {
                this.Table = group.Column.Table;
            }
            else if (group.Table != null)
            {
                this.Table = group.Table;
            }
            else
            {
                this.Table = table;
            }

            if (this.Table.ChildTables.Count > 0 || this.Table.ParentTable != null)
            {
                bool allowChildTables = false;
                if (this.Table.AllowChildTableGroups == InheritableBoolean.Default)
                {
                    allowChildTables = this.Table.GridEX.AllowChildTableGroups;
                }
                else if (this.Table.AllowChildTableGroups == InheritableBoolean.True)
                {
                    allowChildTables = true;
                }
                if (allowChildTables)
                {
                    this.FillTablesCombo();
                    this.grbTable.Visible = true;
                    if (this.Table.ChildTables.Count == 0)
                    {
                        this.cboTables.ReadOnly = true;
                    }
                }
                else
                {
                    this.grbTable.Visible = false;
                }
            }
            else
            {
                this.grbTable.Visible = false;
            }

            this.Column = group.Column;
            // Add any initialization after the InitializeComponent() call.

        }

        protected override void OnLoad(System.EventArgs e)
        {
            base.OnLoad(e);
            //MainQD.MainForm.VisualStyleManager1.AddControl(this, true);
        }

        private GridEXTable Table
        {
            get
            {
                return mTable;
            }
            set
            {
                if (!(mTable == value))
                {
                    mTable = value;
                    this.OnTableChanged();
                }
            }
        }

        private GridEXColumn Column
        {
            get
            {
                return mColumn;
            }
            set
            {
                if (!(value == mColumn))
                {
                    mColumn = value;
                    this.OnColumnChanged();
                }
            }
        }

        private void OnColumnChanged()
        {
            mGroup.Column = mColumn;
            if (mColumn.DataTypeCode == TypeCode.DateTime)
            {
                this.lblGroupInterval.Visible = true;
                this.cboGroupInterval.Visible = true;
            }
            else
            {
                this.lblGroupInterval.Visible = false;
                this.cboGroupInterval.Visible = false;
            }
            if (RefreshData != null)
                RefreshData(this, EventArgs.Empty);

        }

        private void OnTableChanged()
        {
            this.FillColumnsCombo();
        }

        private void FillGroupIntervalCombo()
        {
            this.cboGroupInterval.Items.Add(Janus.Windows.GridEX.GroupInterval.Default);
            this.cboGroupInterval.Items.Add(Janus.Windows.GridEX.GroupInterval.Hour);
            this.cboGroupInterval.Items.Add(Janus.Windows.GridEX.GroupInterval.Minute);
            this.cboGroupInterval.Items.Add(Janus.Windows.GridEX.GroupInterval.Month);
            this.cboGroupInterval.Items.Add(Janus.Windows.GridEX.GroupInterval.Quarter);
            this.cboGroupInterval.Items.Add(Janus.Windows.GridEX.GroupInterval.Second);
            this.cboGroupInterval.Items.Add(Janus.Windows.GridEX.GroupInterval.Year);
            this.cboGroupInterval.Items.Add(Janus.Windows.GridEX.GroupInterval.Text);
            this.cboGroupInterval.Items.Add(Janus.Windows.GridEX.GroupInterval.Value);

            this.cboGroupInterval.SelectedValue = this.mGroup.GroupInterval;
        }

        private void FillColumnsCombo()
        {
            this.cboColumns.Items.Clear();
            foreach (GridEXColumn col in this.Table.Columns)
            {
                if (col.AllowGroup)
                {
                    this.cboColumns.Items.Add(MainQD.GetColumnFriendlyName(col), col);
                }
            }

            this.cboColumns.SelectedValue = mGroup.Column;
            if (cboColumns.SelectedValue == null)
            {
                cboColumns.Text = "Choose a column";
            }
        }

        private void FillTablesCombo()
        {
            this.cboTables.ImageList = mTable.GridEX.ImageList;
            GridEXTable rootTable= mTable;

            while (rootTable.ParentTable != null)
            {
                rootTable = rootTable.ParentTable;
            }

            this.AddChildTables(rootTable, 0);
            this.cboTables.SelectedValue = Table;
  
        }

        private void AddChildTables(GridEXTable table, int indent)
        {
            Janus.Windows.EditControls.UIComboBoxItem item = new Janus.Windows.EditControls.UIComboBoxItem(table.Caption, table);
            item.ImageIndex = table.ImageIndex;
            item.IndentLevel = indent;

            this.cboTables.Items.Add(item);
            foreach (GridEXTable childTable in table.ChildTables)
            {
                this.AddChildTables(childTable, indent + 1);
            }

        }

        private void cboColumns_SelectedItemChanged(object sender, System.EventArgs e)
        {
            this.Column = this.cboColumns.SelectedItem.Value as GridEXColumn;
        }

        private void cboGroupInterval_SelectedItemChanged(object sender, System.EventArgs e)
        {
            this.mGroup.GroupInterval = (GroupInterval)this.cboGroupInterval.SelectedValue;
        }

        private void cboTables_SelectedItemChanged(object sender, System.EventArgs e)
        {
            this.Table = this.cboTables.SelectedItem.Value as GridEXTable;
        }

        private void optAscending_CheckedChanged(object sender, System.EventArgs e)
        {
            if (optAscending.Checked && mGroup != null)
            {
                this.mGroup.SortOrder = Janus.Windows.GridEX.SortOrder.Ascending;
            }
        }

        private void optDescending_CheckedChanged(object sender, System.EventArgs e)
        {
            if (optDescending.Checked && mGroup != null)
            {
                this.mGroup.SortOrder = Janus.Windows.GridEX.SortOrder.Descending;
            }
        }


        #region IEditGroupControl Members

        public event EventHandler RefreshData;

        #endregion
    }


	internal interface IEditGroupControl
	{

        event EventHandler RefreshData;

	}

} //end of root namespace