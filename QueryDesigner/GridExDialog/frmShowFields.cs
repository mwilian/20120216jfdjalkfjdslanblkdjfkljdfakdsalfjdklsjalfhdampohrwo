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
	public partial class frmShowFields
	{

		private ArrayList mAvailableFields;
		private ArrayList mVisibleFields;
		protected override void OnLoad(System.EventArgs e)
		{
            base.OnLoad(e);
            //MainQD.MainForm.VisualStyleManager1.AddControl(this, true);
		}
		public DialogResult ShowDialog(Janus.Windows.GridEX.GridEX grid, Form parent)
		{
			GridEXColumn column = null;
			int i = 0;
			this.lbAvail.DisplayMember = "Caption";
			this.lbVisible.DisplayMember = "Caption";

            for (i = 0; i < grid.RootTable.Columns.Count; i++)
			{
				column = grid.RootTable.Columns.GetColumnInPosition(i);
				if (column.ShowInFieldChooser)
				{
					if (column.Visible)
					{
						AddVisibleField(column, false);
					}
					else
					{
						AddAvailableField(column, false);
					}
				}
			}
			FillAvailableList();
			FillVisibleList();
			return this.ShowDialog(parent);
		}
		private void AddAvailableField(GridEXColumn column, bool refresh)
		{
			if (mAvailableFields == null)
			{
				mAvailableFields = new ArrayList();
			}
			mAvailableFields.Add(column);
			if (refresh)
			{
				FillAvailableList();
			}
		}
		private void FillAvailableList()
		{
			this.lbAvail.Items.Clear();
			if (mAvailableFields != null)
			{
				foreach (GridEXColumn column in mAvailableFields)
				{
					lbAvail.Items.Add(MainQD.GetColumnFriendlyName(column));
				}
			}
			if (lbAvail.Items.Count > 0)
			{
				lbAvail.SelectedIndex = 0;
				btnAdd.Enabled = true;
			}
			else
			{
				btnAdd.Enabled = false;
			}
		}
		private void AddVisibleField(GridEXColumn column, bool refresh)
		{
			if (mVisibleFields == null)
			{
				mVisibleFields = new ArrayList();
			}
			mVisibleFields.Add(column);
			if (refresh)
			{
				FillVisibleList();
			}
		}

		private void FillVisibleList()
		{
			this.lbVisible.Items.Clear();
			if (mVisibleFields != null)
			{
				foreach (GridEXColumn column in mVisibleFields)
				{
					lbVisible.Items.Add(MainQD.GetColumnFriendlyName(column));
				}
			}
			if (lbVisible.Items.Count > 0)
			{
				lbVisible.SelectedIndex = 0;
				btnRemove.Enabled = true;
			}
			else
			{
				btnRemove.Enabled = false;
			}
		}

		private void btnAdd_Click(object sender, System.EventArgs e)
		{
			GridEXColumn column = null;
			int selIndex = 0;
			if (lbAvail.SelectedIndex != -1)
			{
				selIndex = lbAvail.SelectedIndex;
				column = (GridEXColumn)(mAvailableFields[selIndex]);
				mAvailableFields.Remove(column);
				FillAvailableList();
				this.AddVisibleField(column, true);
				lbVisible.SelectedIndex = lbVisible.Items.Count - 1;
				if (selIndex < lbAvail.Items.Count)
				{
					lbAvail.SelectedIndex = selIndex;
				}
				else
				{
					lbAvail.SelectedIndex = lbAvail.Items.Count - 1;
				}
			}
		}

		private void btnRemove_Click(object sender, System.EventArgs e)
		{
			GridEXColumn column = null;
			int selIndex = 0;
			if (lbVisible.SelectedIndex != -1)
			{
				selIndex = lbVisible.SelectedIndex;
				column = (GridEXColumn)(mVisibleFields[selIndex]);
				mVisibleFields.Remove(column);
				FillVisibleList();
				this.AddAvailableField(column, true);
				lbAvail.SelectedIndex = lbAvail.Items.Count - 1;
				if (selIndex < lbVisible.Items.Count)
				{
					lbVisible.SelectedIndex = selIndex;
				}
				else
				{
					lbVisible.SelectedIndex = lbVisible.Items.Count - 1;
				}
			}
		}

		private void lbVisible_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.btnUp.Enabled = (lbVisible.SelectedIndex > 0);
			this.btnDown.Enabled = (lbVisible.SelectedIndex < lbVisible.Items.Count - 1);
		}

		private void btnUp_Click(object sender, System.EventArgs e)
		{
			GridEXColumn column = null;
			int selIndex = lbVisible.SelectedIndex;
			column = (GridEXColumn)(this.mVisibleFields[selIndex]);
			mVisibleFields.Remove(column);
			mVisibleFields.Insert(selIndex - 1, column);
			FillVisibleList();
			lbVisible.SelectedIndex = selIndex - 1;
		}

		private void btnDown_Click(object sender, System.EventArgs e)
		{
			GridEXColumn column = null;
			int selIndex = lbVisible.SelectedIndex;
			column = (GridEXColumn)(this.mVisibleFields[selIndex]);
			mVisibleFields.Remove(column);
			mVisibleFields.Insert(selIndex + 1, column);
			FillVisibleList();
			lbVisible.SelectedIndex = selIndex + 1;
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			int pos = 0;
			if (mAvailableFields != null)
			{
				foreach (GridEXColumn column in mAvailableFields)
				{
					column.Visible = false;
				}
			}
			if (mVisibleFields != null)
			{
				pos = 0;
				foreach (GridEXColumn column in mVisibleFields)
				{
					column.Visible = true;
					column.Position = pos;
					pos = pos + 1;
				}
			}
			this.Close();
		}

		private void lbAvail_DoubleClick(object sender, System.EventArgs e)
		{
			this.btnAdd_Click(null, EventArgs.Empty);
		}

		private void lbVisible_DoubleClick(object sender, System.EventArgs e)
		{
			this.btnRemove_Click(null, EventArgs.Empty);
		}


	}

} //end of root namespace