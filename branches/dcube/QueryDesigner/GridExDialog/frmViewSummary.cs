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
	public partial class frmViewSummary
	{

		private GridEX mGridEX;
		private GridEXView mView;
		protected override void OnLoad(System.EventArgs e)
		{
			base.OnLoad(e);
			NorthwindApp.MainForm.VisualStyleManager1.AddControl(this, true);
		}
		public System.Windows.Forms.DialogResult ShowDialog(GridEX grid, Form parent, GridEXView gridEXView)
		{
			mGridEX = grid;
			mView = gridEXView;
			SetFieldsLabel();
			SetGroupByLabel();
			SetSortLabel();
			SetFormatLabel();
			SetFilterLabel();
			if (grid.RootTable.CellLayoutMode == CellLayoutMode.UseColumnSets)
			{
				this.btnFields.Enabled = false;
			}
			if (mGridEX.View != Janus.Windows.GridEX.View.TableView)
			{
				this.btnGroupBy.Enabled = false;
			}
			return this.ShowDialog(parent);
		}
		private void SetFieldsLabel()
		{
			System.Text.StringBuilder strFields = new System.Text.StringBuilder();
			GridEXColumn column = null;
            for (int i = 0; i < mGridEX.RootTable.Columns.Count; i++)
			{
				column = mGridEX.RootTable.Columns.GetColumnInPosition(i);
				if (column != null && column.Visible)
				{
					if (strFields.Length > 0)
					{
						strFields.Append(", ");
					}
					strFields.Append(NorthwindApp.GetColumnFriendlyName(column));
				}
			}
			this.lblFields.Text = strFields.ToString();
			if (mGridEX.RootTable.CellLayoutMode == CellLayoutMode.UseColumnSets)
			{
				this.btnFields.Enabled = false;
			}
		}
		private void SetGroupByLabel()
		{

			System.Text.StringBuilder strFields = new System.Text.StringBuilder();
			if (mGridEX.RootTable.Groups.Count == 0)
			{
				this.lblGroupBy.Text = "None";
			}
			else
			{
				foreach (GridEXGroup group in mGridEX.RootTable.Groups)
				{
					if (strFields.Length > 0)
					{
						strFields.Append(", ");
					}
					if (group.Column != null)
					{
						strFields.Append(NorthwindApp.GetColumnFriendlyName(group.Column));
					}
					else
					{
						strFields.Append(group.HeaderCaption);
					}
					if (group.SortOrder == Janus.Windows.GridEX.SortOrder.Ascending)
					{
						strFields.Append(" (Ascending)");
					}
					else
					{
						strFields.Append(" (Descending)");
					}
				}

				this.lblGroupBy.Text = strFields.ToString();
			}
		}

		private void SetFilterLabel()
		{
			if (mGridEX.RootTable.FilterCondition == null)
			{
				this.lblFilterBy.Text = "None";
			}
			else
			{
				this.lblFilterBy.Text = mGridEX.RootTable.FilterApplied.ToString();
			}
		}

		private void SetSortLabel()
		{
			System.Text.StringBuilder strFields = new System.Text.StringBuilder();
			if (mGridEX.RootTable.SortKeys.Count == 0)
			{
				this.lblSort.Text = "None";
			}
			else
			{
				foreach (GridEXSortKey sortKey in mGridEX.RootTable.SortKeys)
				{
					if (strFields.Length > 0)
					{
						strFields.Append(", ");
					}
					strFields.Append(NorthwindApp.GetColumnFriendlyName(sortKey.Column));
                    if (sortKey.SortOrder == Janus.Windows.GridEX.SortOrder.Ascending)
					{
						strFields.Append(" (Ascending)");
					}
					else
					{
						strFields.Append(" (Descending)");
					}
				}
				this.lblSort.Text = strFields.ToString();
			}
		}
		private void SetFormatLabel()
		{
            if (mGridEX.View == Janus.Windows.GridEX.View.TableView)
			{
				lblFormat.Text = "Fonts and other Table View settings.";
			}
			else
			{
				lblFormat.Text = "Fonts and other Card View settings.";
			}
		}
		private void btnFields_Click(object sender, System.EventArgs e)
		{
			mView.ShowFieldsDialog();
			SetFieldsLabel();
		}

		private void btnGroupBy_Click(object sender, System.EventArgs e)
		{
			mView.ShowGroupByDialog();
			SetGroupByLabel();
		}

		private void btnSort_Click(object sender, System.EventArgs e)
		{
			mView.ShowSortDialog();
			SetSortLabel();
		}

		private void btnFormat_Click(object sender, System.EventArgs e)
		{
			mView.ShowFormatViewDialog();
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnAutoFormatting_Click(object sender, System.EventArgs e)
		{
			mView.ShowFormatConditionsDialog();
		}

		private void btnFilterBy_Click(object sender, System.EventArgs e)
		{
			mView.ShowFilterDialog();
			this.SetFilterLabel();
		}
	}

} //end of root namespace