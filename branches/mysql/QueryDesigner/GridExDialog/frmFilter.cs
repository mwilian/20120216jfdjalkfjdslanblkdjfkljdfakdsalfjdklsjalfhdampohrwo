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
	public partial class frmFilter
	{

		public DialogResult ShowDialog(GridEX grid, Form parent)
		{

			this.FilterEditor1.SourceControl = grid;
			this.ShowDialog(parent);
			if (this.DialogResult == System.Windows.Forms.DialogResult.OK)
			{
				grid.RootTable.FilterCondition = (IFilterCondition)this.FilterEditor1.FilterCondition;
			}
			return this.DialogResult;
		}

		protected override void OnLoad(System.EventArgs e)
		{
			base.OnLoad(e);
            //NorthwindApp.MainForm.VisualStyleManager1.AddControl(this, true);
		}
	}
} //end of root namespace