namespace QueryDesigner
{
    partial class frmDAView
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
            Janus.Windows.GridEX.GridEXLayout dgvLIST_DAView_Layout_0 = new Janus.Windows.GridEX.GridEXLayout();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDAView));
            this.panel1 = new System.Windows.Forms.Panel();
            this.btRefresh = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.btSave = new System.Windows.Forms.Button();
            this.dgvLIST_DAView = new Janus.Windows.GridEX.GridEX();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvLIST_DAView)).BeginInit();
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
            this.panel1.Size = new System.Drawing.Size(727, 37);
            this.panel1.TabIndex = 17;
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
            this.btCancel.Location = new System.Drawing.Point(625, 3);
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
            this.btSave.Location = new System.Drawing.Point(522, 3);
            this.btSave.Name = "btSave";
            this.btSave.Size = new System.Drawing.Size(97, 32);
            this.btSave.TabIndex = 14;
            this.btSave.Text = "OK";
            this.btSave.UseVisualStyleBackColor = true;
            this.btSave.Click += new System.EventHandler(this.btSave_Click);
            // 
            // dgvLIST_DAView
            // 
            this.dgvLIST_DAView.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.False;
            this.dgvLIST_DAView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvLIST_DAView.FilterMode = Janus.Windows.GridEX.FilterMode.Automatic;
            this.dgvLIST_DAView.FilterRowButtonStyle = Janus.Windows.GridEX.FilterRowButtonStyle.ConditionOperatorDropDown;
            this.dgvLIST_DAView.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            dgvLIST_DAView_Layout_0.IsCurrentLayout = true;
            dgvLIST_DAView_Layout_0.Key = "dgvQDView";
            dgvLIST_DAView_Layout_0.LayoutString = resources.GetString("dgvLIST_DAView_Layout_0.LayoutString");
            this.dgvLIST_DAView.Layouts.AddRange(new Janus.Windows.GridEX.GridEXLayout[] {
            dgvLIST_DAView_Layout_0});
            this.dgvLIST_DAView.Location = new System.Drawing.Point(0, 0);
            this.dgvLIST_DAView.Name = "dgvLIST_DAView";
            this.dgvLIST_DAView.SettingsKey = "dgvLIST_DAView";
            this.dgvLIST_DAView.Size = new System.Drawing.Size(727, 446);
            this.dgvLIST_DAView.TabIndex = 19;
            this.dgvLIST_DAView.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            this.dgvLIST_DAView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.dgvQDView_MouseDoubleClick);
            this.dgvLIST_DAView.RowDoubleClick += new Janus.Windows.GridEX.RowActionEventHandler(this.dgvLIST_DAView_RowDoubleClick);
            this.dgvLIST_DAView.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dgvQDView_KeyUp);
            // 
            // frmDAView
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(727, 483);
            this.Controls.Add(this.dgvLIST_DAView);
            this.Controls.Add(this.panel1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "frmDAView";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Data Access List";
            this.Load += new System.EventHandler(this.Form_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvLIST_DAView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btRefresh;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Button btSave;
        private Janus.Windows.GridEX.GridEX dgvLIST_DAView;
    }
}