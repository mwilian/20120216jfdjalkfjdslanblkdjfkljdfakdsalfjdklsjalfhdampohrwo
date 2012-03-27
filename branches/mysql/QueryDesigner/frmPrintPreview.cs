using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;

namespace QueryDesigner
{
	public partial class frmPrintPreview : System.Windows.Forms.Form
	{

		private Form mOwnerForm;

		protected override void OnLoad(System.EventArgs e)
		{
			base.OnLoad(e);
            //MainQD.MainForm.VisualStyleManager1.AddControl(this, true);
		}
		protected override void OnClosed(System.EventArgs e)
		{
			mOwnerForm.WindowState = this.WindowState;

			if (this.WindowState == FormWindowState.Normal)
			{
				mOwnerForm.Bounds = this.Bounds;
			}
			mOwnerForm.Show();
		}

		public void Show(System.Drawing.Printing.PrintDocument printDocument, Form ownerForm)
		{
			this.Text = "Print Preview - " + printDocument.DocumentName;
			mOwnerForm = ownerForm;
			if (mOwnerForm.WindowState == FormWindowState.Normal)
			{
				Bounds = mOwnerForm.Bounds;
			}
			else
			{
				this.WindowState = mOwnerForm.WindowState;
			}
			mOwnerForm.Hide();
			this.Show();
			this.Update();
			this.PrintPreviewControl1.Document = printDocument;
			this.PrintPreviewControl1.AutoZoom = true;



		}

		private void printPreviewCommands_CommandClick(object sender, Janus.Windows.UI.CommandBars.CommandEventArgs e)
		{
			switch (e.Command.Key)
			{
				case "cmdMoveUp":
					this.PrintPreviewControl1.StartPage = this.PrintPreviewControl1.StartPage - 1;
					break;
				case "cmdMoveDown":
					this.PrintPreviewControl1.StartPage = this.PrintPreviewControl1.StartPage + 1;
					break;
				case "cmdZoom100":
                    this.printPreviewCommands.Commands["cmdOnePage"].Checked = Janus.Windows.UI.InheritableBoolean.False;
					this.printPreviewCommands.Commands["cmdTwoPages"].Checked = Janus.Windows.UI.InheritableBoolean.False;
					this.PrintPreviewControl1.AutoZoom = false;
					this.PrintPreviewControl1.Zoom = 1;
					break;
				case "cmdOnePage":
					this.printPreviewCommands.Commands["cmdZoom100"].Checked = Janus.Windows.UI.InheritableBoolean.False;
                    this.printPreviewCommands.Commands["cmdTwoPages"].Checked = Janus.Windows.UI.InheritableBoolean.False;
					this.PrintPreviewControl1.AutoZoom = true;
					this.PrintPreviewControl1.Rows = 1;
					this.PrintPreviewControl1.Columns = 1;
					break;
				case "cmdTwoPages":
                    this.printPreviewCommands.Commands["cmdZoom100"].Checked = Janus.Windows.UI.InheritableBoolean.False;
                    this.printPreviewCommands.Commands["cmdOnePage"].Checked = Janus.Windows.UI.InheritableBoolean.False;
					this.PrintPreviewControl1.AutoZoom = true;
					this.PrintPreviewControl1.Rows = 1;
					this.PrintPreviewControl1.Columns = 2;
					break;
				case "cmdPageSetup":
					this.PageSetupDialog1.Document = this.PrintPreviewControl1.Document;
					if (this.PageSetupDialog1.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
					{
						System.Drawing.Printing.PrintDocument doc = null;
						doc = this.PrintPreviewControl1.Document;
						this.PrintPreviewControl1.Document = doc;
					}
					break;
				case "cmdPrint":
                    try
                    {
                        this.PrintPreviewControl1.Document.Print();
                    }
                    catch (Exception ex)
                    {
                    }
					this.Close();
					break;
				case "cmdClose":
					this.Close();
					break;
			}
		}


		private void PrintPreviewControl1_StartPageChanged(object sender, System.EventArgs e)
		{

			this.printPreviewCommands.Commands["cmdMoveUp"].IsEnabled = (this.PrintPreviewControl1.StartPage > 0);

		}

        private void frmPrintPreview_Load(object sender, EventArgs e)
        {

        }
	}

} //end of root namespace