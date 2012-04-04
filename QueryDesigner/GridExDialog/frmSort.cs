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
    public partial class frmSort
    {
        protected override void OnLoad(System.EventArgs e)
        {
            base.OnLoad(e);
            //MainQD.MainForm.VisualStyleManager1.AddControl(this, true);
        }
        public DialogResult ShowDialog(GridEX grid, Form parent)
        {
            GridEXSortKey sortKey = null;
            this.FillColumnList(grid.RootTable.Columns, this.cboColumn0);
            this.FillColumnList(grid.RootTable.Columns, this.cboColumn1);
            this.FillColumnList(grid.RootTable.Columns, this.cboColumn2);
            this.FillColumnList(grid.RootTable.Columns, this.cboColumn3);
            if (grid.RootTable.SortKeys.Count == 0)
            {
                SetSortKey(null, true, cboColumn0, optAscending0, optDescending0);
            }
            else
            {
                if (grid.RootTable.SortKeys.Count >= 1)
                {
                    sortKey = grid.RootTable.SortKeys[0];
                    SetSortKey(sortKey.Column, (sortKey.SortOrder == Janus.Windows.GridEX.SortOrder.Ascending), cboColumn0, optAscending0, optDescending0);
                }
                if (grid.RootTable.SortKeys.Count >= 2)
                {
                    sortKey = grid.RootTable.SortKeys[1];
                    SetSortKey(sortKey.Column, (sortKey.SortOrder == Janus.Windows.GridEX.SortOrder.Ascending), cboColumn1, optAscending1, optDescending1);
                }
                if (grid.RootTable.SortKeys.Count >= 3)
                {
                    sortKey = grid.RootTable.SortKeys[2];
                    SetSortKey(sortKey.Column, (sortKey.SortOrder == Janus.Windows.GridEX.SortOrder.Ascending), cboColumn2, optAscending2, optDescending2);
                }
                if (grid.RootTable.SortKeys.Count >= 4)
                {
                    sortKey = grid.RootTable.SortKeys[3];
                    SetSortKey(sortKey.Column, (sortKey.SortOrder == Janus.Windows.GridEX.SortOrder.Ascending), cboColumn3, optAscending3, optDescending3);
                }
            }
            this.ShowDialog(parent);
            if (this.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                grid.RootTable.SortKeys.Clear();
                CreateSortKeys(grid);
            }
            return this.DialogResult;
        }
        private void CreateSortKeys(GridEX grid)
        {
            int sortKeysCount = 0;
            if (cboColumn3.SelectedIndex > 0)
            {
                sortKeysCount = 4;
            }
            else if (cboColumn2.SelectedIndex > 0)
            {
                sortKeysCount = 3;
            }
            else if (cboColumn1.SelectedIndex > 0)
            {
                sortKeysCount = 2;
            }
            else if (cboColumn0.SelectedIndex > 0)
            {
                sortKeysCount = 1;
            }
            else
            {
                sortKeysCount = 0;
            }
            GridEXSortKey[] sortKeys = new GridEXSortKey[sortKeysCount];
            if (sortKeysCount > 0)
            {
                sortKeys[0] = CreateSortKey((GridEXColumn)((object[])cboColumn0.SelectedItem)[1], optAscending0.Checked);
            }
            if (sortKeysCount > 1)
            {
                sortKeys[1] = CreateSortKey((GridEXColumn)((object[])cboColumn1.SelectedItem)[1], optAscending1.Checked);
            }
            if (sortKeysCount > 2)
            {
                sortKeys[2] = CreateSortKey((GridEXColumn)((object[])cboColumn2.SelectedItem)[1], optAscending2.Checked);
            }
            if (sortKeysCount > 3)
            {
                sortKeys[3] = CreateSortKey((GridEXColumn)((object[])cboColumn3.SelectedItem)[1], optAscending3.Checked);
            }
            grid.RootTable.SortKeys.AddRange(sortKeys);
        }
        private GridEXSortKey CreateSortKey(GridEXColumn column, bool ascending)
        {
            GridEXSortKey sortKey = new GridEXSortKey();
            sortKey.Column = column;
            if (!ascending)
            {
                sortKey.SortOrder = Janus.Windows.GridEX.SortOrder.Descending;
            }
            return sortKey;
        }
        private void FillColumnList(GridEXColumnCollection columns, ComboBox combo)
        {
            GridEXColumn column = null;
            int i = 0;

            //combo.DisplayMember = "Name"
            combo.Items.Clear();
            combo.Items.Add(new object[] { "(None)", null });
            for (i = 0; i < columns.Count; i++)
            {
                column = columns[i];
                if (column.AllowSort)
                {
                    combo.Items.Add(new object[] { MainQD.GetColumnFriendlyName(column), column });
                }
            }
        }

        private void SetSortKey(GridEXColumn column, bool ascending, ComboBox combo, RadioButton optAscending, RadioButton optDescending)
        {
            if (column == null)
            {
                combo.SelectedIndex = 0;
            }
            else
            {
                foreach (object item in combo.Items)
                {
                    object[] a = item as object[];
                    if (a.Length == 2)
                        if (a[1] == column)//item.Value
                        {
                            combo.SelectedItem = item;
                            break;
                        }
                }
            }
            if (ascending)
            {
                optAscending.Checked = true;
            }
            else
            {
                optDescending.Checked = true;
            }
        }
        private void cboColumn0_SelectedItemChanged(object sender, System.EventArgs e)
        {
            if (cboColumn0.SelectedIndex == 0)
            {
                this.optAscending0.Enabled = false;
                this.optDescending0.Enabled = false;
                cboColumn1.SelectedIndex = 0;
                cboColumn1.Enabled = false;
            }
            else
            {
                this.optAscending0.Enabled = true;
                this.optDescending0.Enabled = true;
                cboColumn1.Enabled = true;
                if (cboColumn1.SelectedIndex == -1)
                {
                    cboColumn1.SelectedIndex = 0;
                }
            }
        }

        private void cboColumn1_SelectedItemChanged(object sender, System.EventArgs e)
        {
            if (cboColumn1.SelectedIndex == 0)
            {
                this.optAscending1.Enabled = false;
                this.optDescending1.Enabled = false;
                cboColumn2.SelectedIndex = 0;
                cboColumn2.Enabled = false;
            }
            else
            {
                this.optAscending1.Enabled = true;
                this.optDescending1.Enabled = true;
                cboColumn2.Enabled = true;
                if (cboColumn2.SelectedIndex == -1)
                {
                    cboColumn2.SelectedIndex = 0;
                }
            }
        }

        private void cboColumn2_SelectedItemChanged(object sender, System.EventArgs e)
        {
            if (cboColumn2.SelectedIndex == 0)
            {
                this.optAscending2.Enabled = false;
                this.optDescending2.Enabled = false;
                cboColumn3.SelectedIndex = 0;
                cboColumn3.Enabled = false;
            }
            else
            {
                this.optAscending2.Enabled = true;
                this.optDescending2.Enabled = true;
                cboColumn3.Enabled = true;
                if (cboColumn3.SelectedIndex == -1)
                {
                    cboColumn3.SelectedIndex = 0;
                }
            }
        }

        private void cboColumn3_SelectedItemChanged(object sender, System.EventArgs e)
        {
            if (cboColumn3.SelectedIndex == 0)
            {
                this.optAscending3.Enabled = false;
                this.optDescending3.Enabled = false;
            }
            else
            {
                this.optAscending3.Enabled = true;
                this.optDescending3.Enabled = true;
            }
        }

        private void btnClear_Click(object sender, System.EventArgs e)
        {
            cboColumn0.SelectedIndex = 0;
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

    }

} //end of root namespace