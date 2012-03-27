using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using System.Threading;

using FlexCel.Render;

namespace QueryDesigner
{
	/// <summary>
	/// Summary description for PdfProgressDialog.
	/// </summary>
	public class PdfProgressDialog : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.StatusBar statusBar1;
		private System.Windows.Forms.StatusBarPanel statusBarPanel1;
		private System.Windows.Forms.StatusBarPanel statusBarPanelTime;
		private System.Windows.Forms.Label labelPages;
        private System.Timers.Timer timer1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public PdfProgressDialog()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.btnCancel = new System.Windows.Forms.Button();
            this.statusBar1 = new System.Windows.Forms.StatusBar();
            this.statusBarPanel1 = new System.Windows.Forms.StatusBarPanel();
            this.statusBarPanelTime = new System.Windows.Forms.StatusBarPanel();
            this.labelPages = new System.Windows.Forms.Label();
            this.timer1 = new System.Timers.Timer();
            ((System.ComponentModel.ISupportInitialize)(this.statusBarPanel1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.statusBarPanelTime)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.timer1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(184, 64);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Cancel";
            // 
            // statusBar1
            // 
            this.statusBar1.Location = new System.Drawing.Point(0, 100);
            this.statusBar1.Name = "statusBar1";
            this.statusBar1.Panels.AddRange(new System.Windows.Forms.StatusBarPanel[] {
            this.statusBarPanel1,
            this.statusBarPanelTime});
            this.statusBar1.ShowPanels = true;
            this.statusBar1.Size = new System.Drawing.Size(448, 22);
            this.statusBar1.TabIndex = 1;
            // 
            // statusBarPanel1
            // 
            this.statusBarPanel1.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.None;
            this.statusBarPanel1.Name = "statusBarPanel1";
            this.statusBarPanel1.Text = "Elapsed Time:";
            this.statusBarPanel1.Width = 80;
            // 
            // statusBarPanelTime
            // 
            this.statusBarPanelTime.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.None;
            this.statusBarPanelTime.Name = "statusBarPanelTime";
            this.statusBarPanelTime.Text = "0:00";
            // 
            // labelPages
            // 
            this.labelPages.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.labelPages.Location = new System.Drawing.Point(16, 16);
            this.labelPages.Name = "labelPages";
            this.labelPages.Size = new System.Drawing.Size(408, 16);
            this.labelPages.TabIndex = 2;
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.SynchronizingObject = this;
            this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Elapsed);
            // 
            // PdfProgressDialog
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(448, 122);
            this.ControlBox = false;
            this.Controls.Add(this.labelPages);
            this.Controls.Add(this.statusBar1);
            this.Controls.Add(this.btnCancel);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PdfProgressDialog";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Please wait...";
            this.Closed += new System.EventHandler(this.PdfProgressDialog_Closed);
            ((System.ComponentModel.ISupportInitialize)(this.statusBarPanel1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.statusBarPanelTime)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.timer1)).EndInit();
            this.ResumeLayout(false);

		}
		#endregion


		private DateTime StartTime;
		private Thread RunningThread;
		private FlexCelPdfExport PdfExport;

		private void timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
		{
			UpdateStatus();		
		}

		public void ShowProgress(Thread aRunningThread, FlexCelPdfExport aPdfExport)
		{
			RunningThread = aRunningThread;

			if (!RunningThread.IsAlive) { DialogResult = DialogResult.OK; return; }
			timer1.Enabled = true;
			StartTime = DateTime.Now;
			PdfExport = aPdfExport;
			ShowDialog();
		}

		private void UpdateStatus()
		{
			TimeSpan ts = DateTime.Now - StartTime;
			string hours = ts.Hours == 0 ? "" : ts.Hours.ToString("00") + ":";
			statusBarPanelTime.Text = hours + ts.Minutes.ToString("00") + ":" + ts.Seconds.ToString("00");

			if (!RunningThread.IsAlive) DialogResult = DialogResult.OK;

			if (PdfExport.Progress.TotalPage > 0) labelPages.Text = String.Format("Generating Page {0} of {1}", PdfExport.Progress.Page, PdfExport.Progress.TotalPage);
		}

		private void PdfProgressDialog_Closed(object sender, System.EventArgs e)
		{
			timer1.Enabled = false;
		}


	}
}
