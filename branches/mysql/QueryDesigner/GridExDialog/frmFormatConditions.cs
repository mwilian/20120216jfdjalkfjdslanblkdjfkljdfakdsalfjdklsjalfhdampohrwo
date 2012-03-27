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
	public partial class frmFormatConditions
	{

		private bool mTwoValues;
		private Janus.Windows.GridEX.GridEX mGridEX;
		private System.Collections.ArrayList tempConditions;
		private GridEXFormatCondition mActiveCondition;
		protected override void OnLoad(System.EventArgs e)
		{
			base.OnLoad(e);
            //MainQD.MainForm.VisualStyleManager1.AddControl(this, true);
		}
		private GridEXFormatCondition ActiveCondition
		{
			get
			{
				return mActiveCondition;
			}
			set
			{
				if (value != mActiveCondition)
				{
					if (mActiveCondition != null)
					{
						this.GetConditionValues();
					}
					mActiveCondition = value;
					OnActiveConditionChanged();
				}
			}
		}

		private void PopulateCombos()
		{
			cboFields.DataSource = mGridEX.RootTable.Columns;
			cboFields.DisplayMember = "DataMember";

			Array values = System.Enum.GetValues(Janus.Windows.GridEX.ConditionOperator.EndsWith.GetType());
			foreach (object value in values)
			{
				cboCondition.Items.Add(value);
			}

			this.chkBold.IndeterminatedValue = Janus.Windows.GridEX.TriState.Empty;
			this.chkItalic.IndeterminatedValue = Janus.Windows.GridEX.TriState.Empty;
			this.chkStrikeout.IndeterminatedValue = Janus.Windows.GridEX.TriState.Empty;
			this.chkUnderline.IndeterminatedValue = Janus.Windows.GridEX.TriState.Empty;

			this.chkBold.CheckedValue = Janus.Windows.GridEX.TriState.True;
			this.chkItalic.CheckedValue = Janus.Windows.GridEX.TriState.True;
			this.chkStrikeout.CheckedValue = Janus.Windows.GridEX.TriState.True;
			this.chkUnderline.CheckedValue = Janus.Windows.GridEX.TriState.True;

			this.chkBold.UncheckedValue = Janus.Windows.GridEX.TriState.False;
			this.chkItalic.UncheckedValue = Janus.Windows.GridEX.TriState.False;
			this.chkStrikeout.UncheckedValue = Janus.Windows.GridEX.TriState.False;
			this.chkUnderline.UncheckedValue = Janus.Windows.GridEX.TriState.False;

		}

		public void ShowDialog(Janus.Windows.GridEX.GridEX gridEx, System.Windows.Forms.Form parentForm)
		{
			mGridEX = gridEx;
			this.PopulateCombos();

			tempConditions = new System.Collections.ArrayList();
			foreach (GridEXFormatCondition cond in this.mGridEX.RootTable.FormatConditions)
			{
				GridEXFormatCondition cloneCond = cond.Clone();
				tempConditions.Add(cloneCond);
			}

			this.jsgConditions.RootTable.Columns["clmName"].DataMember = "Key";
			this.jsgConditions.RootTable.Columns["clmEnabled"].DataMember = "Enabled";
			this.jsgConditions.SetDataBinding(tempConditions, "");
			OnActiveConditionChanged();
			this.ShowDialog(parentForm);
		}

		private bool GetConditionValues()
		{
			if (mActiveCondition != null)
			{
				if (this.cboFields.SelectedItem == null)
				{
					MessageBox.Show("Select the field for the condition", MainQD.MessageCaption);
					this.cboFields.Focus();
					return false;
				}
				GridEXColumn column = (GridEXColumn)this.cboFields.SelectedItem.Value;
				try
				{
					object value1 = Convert.ChangeType(txtValue1.Text, column.Type);
					mActiveCondition.Value1 = value1;
				}
				catch
				{
					MessageBox.Show("Value 1 is not valid", MainQD.MessageCaption);
					txtValue1.Focus();
					txtValue1.SelectionStart = txtValue1.Text.Length;
					return false;
				}

				if (mTwoValues)
				{

					try
					{
						object value2 = Convert.ChangeType(txtValue2.Text, column.Type);
						mActiveCondition.Value2 = value2;
					}
					catch
					{
						MessageBox.Show("Value 2 is not valid", MainQD.MessageCaption);
						txtValue2.SelectionStart = 0;
						return false;

					}
				}
				else
				{

					mActiveCondition.Value2 = null;
				}
				mActiveCondition.Key = txtConditionName.Text;
				mActiveCondition.Column = column;
				mActiveCondition.ConditionOperator = (ConditionOperator)this.cboCondition.SelectedItem.Value;
				mActiveCondition.FormatStyle.BackColor = this.btnBackColor.SelectedColor;
				mActiveCondition.FormatStyle.ForeColor = this.btnForeColor.SelectedColor;
				mActiveCondition.FormatStyle.FontBold = (TriState)this.chkBold.BindableValue;
				mActiveCondition.FormatStyle.FontItalic = (TriState)this.chkItalic.BindableValue;
				mActiveCondition.FormatStyle.FontUnderline = (TriState)this.chkUnderline.BindableValue;
				mActiveCondition.FormatStyle.FontStrikeout = (TriState)this.chkStrikeout.BindableValue;
			}
			return true;
		}

		private void OnActiveConditionChanged()
		{
			if (mActiveCondition == null)
			{
				this.excConditionName.Group.Enabled = false;
				this.excConditionCriteria.Group.Enabled = false;
				this.excAppearance.Group.Enabled = false;
				this.txtConditionName.Text = "";
				this.cboFields.SelectedItem = null;
				this.cboCondition.SelectedValue = null;
				this.txtValue1.Text = "";
				this.txtValue2.Text = "";
				this.btnBackColor.SelectedColor = Color.Empty;
				this.btnForeColor.SelectedColor = Color.Empty;
				this.chkBold.CheckState = CheckState.Indeterminate;
				this.chkItalic.CheckState = CheckState.Indeterminate;
				this.chkUnderline.CheckState = CheckState.Indeterminate;
				this.chkStrikeout.CheckState = CheckState.Indeterminate;
			}
			else
			{
				this.excConditionName.Group.Enabled = true;
				this.excConditionCriteria.Group.Enabled = true;
				this.excAppearance.Group.Enabled = true;
				this.txtConditionName.Text = mActiveCondition.Key;
				if (mActiveCondition.Column == null)
				{
					this.cboFields.SelectedIndex = -1;
				}
				else
				{
					this.cboFields.SelectedValue = mActiveCondition.Column;
				}
				this.cboCondition.SelectedValue = mActiveCondition.ConditionOperator;
				this.txtValue1.Text = mActiveCondition.Value1 + "";
				this.txtValue2.Text = mActiveCondition.Value2 + "";
				this.btnBackColor.SelectedColor = mActiveCondition.FormatStyle.BackColor;
				this.btnForeColor.SelectedColor = mActiveCondition.FormatStyle.ForeColor;
				chkBold.BindableValue = mActiveCondition.FormatStyle.FontBold;
				chkItalic.BindableValue = mActiveCondition.FormatStyle.FontItalic;
				chkUnderline.BindableValue = mActiveCondition.FormatStyle.FontUnderline;
				chkStrikeout.BindableValue = mActiveCondition.FormatStyle.FontStrikeout;
			}
		}

		private void cboCondition_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (cboCondition.SelectedValue != null)
			{
				switch ((ConditionOperator)cboCondition.SelectedValue)
				{
					case ConditionOperator.Between:
					case ConditionOperator.NotBetween:
						txtValue2.Enabled = true;
						lblValue2.Enabled = true;
						mTwoValues = true;
						break;
					default:
						txtValue2.Text = "";
						txtValue2.Enabled = false;
						lblValue2.Enabled = false;
						mTwoValues = false;
						break;
				}
			}
		}

		private void btnMoveUp_Click(object sender, System.EventArgs e)
		{
			int index = this.jsgConditions.Row;
			object condition = this.tempConditions[index];
			this.tempConditions.Remove(condition);
			this.tempConditions.Insert(index - 1, condition);
			this.jsgConditions.Refetch();
			this.jsgConditions.Row = index - 1;
			this.jsgConditions.Focus();
		}

		private void btnMoveDown_Click(object sender, System.EventArgs e)
		{
			int index = this.jsgConditions.Row;
			object condition = this.tempConditions[index];
			this.tempConditions.Remove(condition);
			this.tempConditions.Insert(index + 1, condition);
			this.jsgConditions.Refetch();
			this.jsgConditions.Row = index + 1;
			this.jsgConditions.Focus();

		}

		private void btnNew_Click(object sender, System.EventArgs e)
		{
			if (this.GetConditionValues())
			{
				GridEXFormatCondition condition = new GridEXFormatCondition();
				condition.Key = "Untitled";
				this.tempConditions.Add(condition);
				this.jsgConditions.Refetch();
				this.jsgConditions.Row = this.jsgConditions.RecordCount - 1;
				this.jsgConditions.Focus();
			}
		}

		private void btnDelete_Click(object sender, System.EventArgs e)
		{

			this.jsgConditions.AllowDelete = Janus.Windows.GridEX.InheritableBoolean.True;
			mActiveCondition = null;
			OnActiveConditionChanged();
			this.jsgConditions.Delete();
			this.jsgConditions.AllowDelete = Janus.Windows.GridEX.InheritableBoolean.False;

		}

		private void txtConditionName_TextChanged(object sender, System.EventArgs e)
		{
			if (mActiveCondition != null)
			{
				this.ActiveCondition.Key = txtConditionName.Text;
				this.jsgConditions.Refresh();
			}
		}

		private void jsgConditions_UpdatingRecord(object sender, System.ComponentModel.CancelEventArgs e)
		{
			Janus.Windows.GridEX.GridEXRow row = null;
			row = jsgConditions.GetRow();
			this.txtConditionName.Text = (string)row.Cells["clmName"].Value;

		}

		private void jsgConditions_CurrentCellChanging(object sender, Janus.Windows.GridEX.CurrentCellChangingEventArgs e)
		{
			if (e.Row != null)
			{
				if (this.jsgConditions.Row >= 0 &&  (e.Row.Position != this.jsgConditions.Row))
				{
					if (! this.GetConditionValues())
					{
						e.Cancel = true;
					}
				}
			}
		}

		private void jsgConditions_SelectionChanged(object sender, System.EventArgs e)
		{
			Janus.Windows.GridEX.GridEXRow currentRow = this.jsgConditions.GetRow();
			if (currentRow != null)
			{
				ActiveCondition = (GridEXFormatCondition)currentRow.DataRow;
				if (currentRow.Position == 0)
				{
					this.btnMoveUp.Enabled = false;
				}
				else
				{
					this.btnMoveUp.Enabled = true;
				}
				if (currentRow.Position < this.jsgConditions.RowCount - 1)
				{
					this.btnMoveDown.Enabled = true;
				}
				else
				{
					this.btnMoveDown.Enabled = false;
				}

			}
			else
			{
				this.btnMoveDown.Enabled = false;
				this.btnMoveUp.Enabled = false;
				ActiveCondition = null;
			}
		}


		private void jsgConditions_FormattingRow(object sender, Janus.Windows.GridEX.RowLoadEventArgs e)
		{
			if (e.Row.RowType == Janus.Windows.GridEX.RowType.Record)
			{
				GridEXFormatCondition formatCondition = (GridEXFormatCondition)e.Row.DataRow;
				e.Row.RowStyle = new GridEXFormatStyle(formatCondition.FormatStyle);
			}
		}

		private void FormatsForm_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.DialogResult == System.Windows.Forms.DialogResult.OK)
			{
				if (GetConditionValues())
				{
					mGridEX.RootTable.FormatConditions.Clear();
					foreach (GridEXFormatCondition condition in tempConditions)
					{
						mGridEX.RootTable.FormatConditions.Add(condition);
					}
				}
				else
				{
					e.Cancel = true;
				}
			}
		}

	}

} //end of root namespace