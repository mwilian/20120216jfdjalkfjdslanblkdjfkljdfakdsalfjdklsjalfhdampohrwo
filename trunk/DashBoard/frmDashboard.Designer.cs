namespace DashBoard
{
    partial class frmDashboard
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDashboard));
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.createGadgetToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ucDashboard = new DashBoard.Control.ucDashboard();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.createGadgetToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(146, 26);
            // 
            // createGadgetToolStripMenuItem
            // 
            this.createGadgetToolStripMenuItem.Name = "createGadgetToolStripMenuItem";
            this.createGadgetToolStripMenuItem.Size = new System.Drawing.Size(145, 22);
            this.createGadgetToolStripMenuItem.Text = "Create Gadget";
            this.createGadgetToolStripMenuItem.Click += new System.EventHandler(this.createGadgetToolStripMenuItem_Click);
            // 
            // ucDashboard
            // 
            this.ucDashboard.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ucDashboard.Location = new System.Drawing.Point(0, 0);
            this.ucDashboard.Name = "ucDashboard";
            this.ucDashboard.Size = new System.Drawing.Size(770, 398);
            this.ucDashboard.TabIndex = 5;
            // 
            // frmDashboard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(770, 398);
            this.ContextMenuStrip = this.contextMenuStrip1;
            this.Controls.Add(this.ucDashboard);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmDashboard";
            this.Text = "Dashboard";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem createGadgetToolStripMenuItem;
        private DashBoard.Control.ucDashboard ucDashboard;
    }
}

