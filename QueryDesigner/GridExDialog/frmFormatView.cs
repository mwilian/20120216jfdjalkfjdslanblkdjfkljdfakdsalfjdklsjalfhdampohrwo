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
	public partial class frmFormatView
	{

		private Font mHeaderFont;
		private Font mRowsFont;
		protected override void OnLoad(System.EventArgs e)
		{
			base.OnLoad(e);
            //MainQD.MainForm.VisualStyleManager1.AddControl(this, true);
		}
		public System.Windows.Forms.DialogResult ShowDialog(GridEX grid, Form parent)
		{
			if (grid.HeaderFormatStyle.Font == null)
			{
				HeaderFont = (Font)grid.Font.Clone();
			}
			else
			{
				HeaderFont = (Font)grid.HeaderFormatStyle.Font.Clone();
			}
			if (grid.RowFormatStyle.Font == null)
			{
				RowsFont = (Font)grid.Font.Clone();
			}
			else
			{
				RowsFont = (Font)grid.RowFormatStyle.Font.Clone();
			}
			if (grid.AllowAddNew == InheritableBoolean.True)
			{
				this.chkAllowAddNew.Checked = true;
			}
			if (grid.AllowEdit == InheritableBoolean.True)
			{
				this.chkAllowEdit.Checked = true;
			}
			if (grid.ColumnAutoResize)
			{
				this.chkAutoSize.Checked = true;
			}

			this.cboGridlineStyle.Items.Add("No grid lines");
			this.cboGridlineStyle.Items.Add("Small dots");
			this.cboGridlineStyle.Items.Add("Solid");
			if (grid.GridLines == GridLines.None)
			{
				this.cboGridlineStyle.SelectedIndex = 0;
			}
			else
			{
				if (grid.GridLineStyle == GridLineStyle.SmallDots)
				{
					this.cboGridlineStyle.SelectedIndex = 1;
				}
				else
				{
					this.cboGridlineStyle.SelectedIndex = 2;
				}
			}
			if (grid.GroupRowFormatStyle.BackColor.Equals(SystemColors.Control))
			{
				this.chkShadeGroupHeaders.Checked = true;
			}
			this.ShowDialog(parent);

			if (this.DialogResult == System.Windows.Forms.DialogResult.OK)
			{
				if (mHeaderFont.Equals(grid.Font))
				{
					grid.HeaderFormatStyle.Font = null;
				}
				else
				{
					grid.HeaderFormatStyle.Font = mHeaderFont;
				}
				if (mRowsFont.Equals(grid.Font))
				{
					grid.RowFormatStyle.Font = null;
				}
				else
				{
					grid.RowFormatStyle.Font = mRowsFont;
				}
				if (this.chkAllowAddNew.Checked)
				{
					grid.AllowAddNew = InheritableBoolean.True;
				}
				else
				{
					grid.AllowAddNew = InheritableBoolean.False;
				}
				if (this.chkAllowEdit.Checked)
				{
					grid.AllowEdit = InheritableBoolean.True;
				}
				else
				{
					grid.AllowEdit = InheritableBoolean.False;
				}
				grid.ColumnAutoResize = this.chkAutoSize.Checked;
				switch (this.cboGridlineStyle.SelectedIndex)
				{
					case 0:
						grid.GridLines = GridLines.None;
						break;
					case 1:
						grid.GridLines = GridLines.Both;
						grid.GridLineStyle = GridLineStyle.SmallDots;
						break;
					case 2:
						grid.GridLines = GridLines.Both;
						grid.GridLineStyle = GridLineStyle.Solid;
						break;
				}
				if (this.chkShadeGroupHeaders.Checked)
				{
					grid.ThemedAreas = grid.ThemedAreas | ThemedArea.GroupRows;
					grid.GroupRowFormatStyle.BackColor = SystemColors.Control;
				}
				else
				{
					grid.ThemedAreas = grid.ThemedAreas ^ ThemedArea.GroupRows;
					grid.GroupRowFormatStyle.BackColor = SystemColors.Window;
				}
			}
			return this.DialogResult;
		}
		private Font HeaderFont
		{
			get
			{
				return mHeaderFont;
			}
			set
			{
				if (value == null)
				{
					mHeaderFont = null;
					this.lblHeaderFont.Text = "";
				}
				else
				{
					if (! value.Equals(mHeaderFont))
					{
						mHeaderFont = value;
						this.lblHeaderFont.Font = new Font(mHeaderFont.Name, (float)(this.lblHeaderFont.Font.Size), mHeaderFont.Style);
						this.lblHeaderFont.Text = mHeaderFont.SizeInPoints.ToString() + " pt. " + mHeaderFont.Name;
					}
				}
			}
		}
		private Font RowsFont
		{
			get
			{
				return mRowsFont;
			}
			set
			{
				if (value == null)
				{
					mRowsFont = null;
					this.lblRowsFont.Text = "";
				}
				else
				{
					if (! value.Equals(mRowsFont))
					{
						mRowsFont = value;
						this.lblRowsFont.Font = new Font(mRowsFont.Name, (float)(this.lblRowsFont.Font.Size), mRowsFont.Style);
						this.lblRowsFont.Text = mRowsFont.SizeInPoints.ToString() + " pt. " + mRowsFont.Name;
					}
				}
			}
		}
		private void btnHeaderFont_Click(object sender, System.EventArgs e)
		{
			this.FontDialog1.Font = mHeaderFont;
			this.FontDialog1.ShowColor = false;
			if (this.FontDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				this.HeaderFont = this.FontDialog1.Font;
			}
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnRowsFont_Click(object sender, System.EventArgs e)
		{
			this.FontDialog1.Font = mRowsFont;
			this.FontDialog1.ShowColor = false;
			if (this.FontDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				this.RowsFont = this.FontDialog1.Font;
			}
		}


	}

} //end of root namespace