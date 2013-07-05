namespace dCube
{
    partial class Form_View
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
            Janus.Windows.GridEX.GridEXLayout dgvQDView_Layout_0 = new Janus.Windows.GridEX.GridEXLayout();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_View));
            this.panel1 = new System.Windows.Forms.Panel();
            this.btRefresh = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.btSave = new System.Windows.Forms.Button();
            this.dgvQDView = new Janus.Windows.GridEX.GridEX();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvQDView)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.btRefresh);
            this.panel1.Controls.Add(this.btCancel);
            this.panel1.Controls.Add(this.btSave);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 446);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(714, 37);
            this.panel1.TabIndex = 2;
            // 
            // btRefresh
            // 
            this.btRefresh.Location = new System.Drawing.Point(3, 3);
            this.btRefresh.Name = "btRefresh";
            this.btRefresh.Size = new System.Drawing.Size(97, 32);
            this.btRefresh.TabIndex = 13;
            this.btRefresh.Text = "Refresh";
            this.btRefresh.UseVisualStyleBackColor = true;
            this.btRefresh.Click += new System.EventHandler(this.btnReresh_Click);
            // 
            // btCancel
            // 
            this.btCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btCancel.Location = new System.Drawing.Point(612, 3);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(97, 32);
            this.btCancel.TabIndex = 15;
            this.btCancel.Text = "Cancel";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // btSave
            // 
            this.btSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btSave.Location = new System.Drawing.Point(509, 3);
            this.btSave.Name = "btSave";
            this.btSave.Size = new System.Drawing.Size(97, 32);
            this.btSave.TabIndex = 14;
            this.btSave.Text = "OK";
            this.btSave.UseVisualStyleBackColor = true;
            this.btSave.Click += new System.EventHandler(this.btSave_Click);
            // 
            // dgvQDView
            // 
            this.dgvQDView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvQDView.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dgvQDView_Layout_0.IsCurrentLayout = true;
            dgvQDView_Layout_0.Key = "dgvQDView";
            dgvQDView_Layout_0.LayoutString = resources.GetString("dgvQDView_Layout_0.LayoutString");
            this.dgvQDView.Layouts.AddRange(new Janus.Windows.GridEX.GridEXLayout[] {
            dgvQDView_Layout_0});
            this.dgvQDView.Location = new System.Drawing.Point(0, 0);
            this.dgvQDView.Name = "dgvQDView";
            this.dgvQDView.SettingsKey = "dgvQDView";
            this.dgvQDView.Size = new System.Drawing.Size(714, 446);
            this.dgvQDView.TabIndex = 1;
            this.dgvQDView.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            this.dgvQDView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.dgvQDView_MouseDoubleClick);
            this.dgvQDView.RowDoubleClick += new Janus.Windows.GridEX.RowActionEventHandler(this.dgvQDView_RowDoubleClick);
            this.dgvQDView.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dgvQDView_KeyUp);
            // 
            // Form_View
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(714, 483);
            this.Controls.Add(this.dgvQDView);
            this.Controls.Add(this.panel1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "Form_View";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Query View";
            this.Load += new System.EventHandler(this.Form_View_Load);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form_View_FormClosed);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvQDView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btRefresh;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Button btSave;
        private Janus.Windows.GridEX.GridEX dgvQDView;
    }
}