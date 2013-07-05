namespace dCube
{
    partial class frmImportDefView
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
            Janus.Windows.GridEX.GridEXLayout dgvQDADDView_Layout_0 = new Janus.Windows.GridEX.GridEXLayout();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmImportDefView));
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.dgvQDADDView = new Janus.Windows.GridEX.GridEX();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvQDADDView)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.btnRefresh);
            this.panel1.Controls.Add(this.btnOK);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 461);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(711, 41);
            this.panel1.TabIndex = 0;
            // 
            // btnRefresh
            // 
            this.btnRefresh.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnRefresh.Location = new System.Drawing.Point(11, 2);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(75, 34);
            this.btnRefresh.TabIndex = 2;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.Location = new System.Drawing.Point(541, 2);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 34);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(622, 2);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 34);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // dgvQDADDView
            // 
            this.dgvQDADDView.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.False;
            this.dgvQDADDView.ColumnAutoSizeMode = Janus.Windows.GridEX.ColumnAutoSizeMode.DisplayedCellsAndHeader;
            this.dgvQDADDView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvQDADDView.FilterMode = Janus.Windows.GridEX.FilterMode.Automatic;
            this.dgvQDADDView.FilterRowButtonStyle = Janus.Windows.GridEX.FilterRowButtonStyle.ConditionOperatorDropDown;
            this.dgvQDADDView.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            dgvQDADDView_Layout_0.IsCurrentLayout = true;
            dgvQDADDView_Layout_0.Key = "dgvDatabaseView";
            dgvQDADDView_Layout_0.LayoutString = resources.GetString("dgvQDADDView_Layout_0.LayoutString");
            this.dgvQDADDView.Layouts.AddRange(new Janus.Windows.GridEX.GridEXLayout[] {
            dgvQDADDView_Layout_0});
            this.dgvQDADDView.Location = new System.Drawing.Point(0, 0);
            this.dgvQDADDView.Name = "dgvQDADDView";
            this.dgvQDADDView.SettingsKey = "dgvImportDefView";
            this.dgvQDADDView.Size = new System.Drawing.Size(711, 461);
            this.dgvQDADDView.TabIndex = 18;
            this.dgvQDADDView.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            this.dgvQDADDView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.dgvQDADDView_MouseDoubleClick);
            this.dgvQDADDView.RowDoubleClick += new Janus.Windows.GridEX.RowActionEventHandler(this.dgvQDADDView_RowDoubleClick);
            this.dgvQDADDView.FilterApplied += new System.EventHandler(this.dgvQDADDView_FilterApplied);
            this.dgvQDADDView.GroupsChanged += new Janus.Windows.GridEX.GroupsChangedEventHandler(this.dgvQDADDView_GroupsChanged);
            this.dgvQDADDView.SizingColumn += new Janus.Windows.GridEX.SizingColumnEventHandler(this.dgvQDADDView_SizingColumn);
            this.dgvQDADDView.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dgvQDADDView_KeyUp);
            // 
            // frmImportDefView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(711, 502);
            this.Controls.Add(this.dgvQDADDView);
            this.Controls.Add(this.panel1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "frmImportDefView";
            this.Text = "_Name";
            this.Load += new System.EventHandler(this.frmLookup_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvQDADDView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private Janus.Windows.GridEX.GridEX dgvQDADDView;
        private System.Windows.Forms.Button btnRefresh;

    }
}