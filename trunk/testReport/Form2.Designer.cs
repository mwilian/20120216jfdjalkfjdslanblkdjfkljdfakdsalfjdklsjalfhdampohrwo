namespace testReport
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tvcReportViewer1 = new ReportDLL.TVCReportViewer();
            this.SuspendLayout();
            // 
            // tvcReportViewer1
            // 
            this.tvcReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvcReportViewer1.Location = new System.Drawing.Point(0, 0);
            this.tvcReportViewer1.Name = "tvcReportViewer1";
            this.tvcReportViewer1.ReportSource = null;
            this.tvcReportViewer1.Size = new System.Drawing.Size(895, 428);
            this.tvcReportViewer1.TabIndex = 0;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(895, 428);
            this.Controls.Add(this.tvcReportViewer1);
            this.Name = "Form2";
            this.Text = "Form2";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private ReportDLL.TVCReportViewer tvcReportViewer1;
    }
}