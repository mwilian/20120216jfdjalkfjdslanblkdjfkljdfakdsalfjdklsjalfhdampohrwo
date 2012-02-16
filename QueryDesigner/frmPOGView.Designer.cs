namespace QueryDesigner
{
    partial class frmPOGView
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
            Janus.Windows.GridEX.GridEXLayout dgvPOGView_Layout_0 = new Janus.Windows.GridEX.GridEXLayout();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPOGView));
            this.panel1 = new System.Windows.Forms.Panel();
            this.btRefresh = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.btSave = new System.Windows.Forms.Button();
            this.dgvPOGView = new Janus.Windows.GridEX.GridEX();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPOGView)).BeginInit();
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
            this.panel1.Size = new System.Drawing.Size(520, 37);
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
            this.btCancel.Location = new System.Drawing.Point(418, 3);
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
            this.btSave.Location = new System.Drawing.Point(315, 3);
            this.btSave.Name = "btSave";
            this.btSave.Size = new System.Drawing.Size(97, 32);
            this.btSave.TabIndex = 14;
            this.btSave.Text = "OK";
            this.btSave.UseVisualStyleBackColor = true;
            this.btSave.Click += new System.EventHandler(this.btSave_Click);
            // 
            // dgvPOGView
            // 
            this.dgvPOGView.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.False;
            this.dgvPOGView.ColumnAutoSizeMode = Janus.Windows.GridEX.ColumnAutoSizeMode.DisplayedCellsAndHeader;
            this.dgvPOGView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvPOGView.FilterMode = Janus.Windows.GridEX.FilterMode.Automatic;
            this.dgvPOGView.FilterRowButtonStyle = Janus.Windows.GridEX.FilterRowButtonStyle.ConditionOperatorDropDown;
            this.dgvPOGView.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            dgvPOGView_Layout_0.IsCurrentLayout = true;
            dgvPOGView_Layout_0.Key = "dgvQDView";
            dgvPOGView_Layout_0.LayoutString = resources.GetString("dgvPOGView_Layout_0.LayoutString");
            this.dgvPOGView.Layouts.AddRange(new Janus.Windows.GridEX.GridEXLayout[] {
            dgvPOGView_Layout_0});
            this.dgvPOGView.Location = new System.Drawing.Point(0, 0);
            this.dgvPOGView.Name = "dgvPOGView";
            this.dgvPOGView.SettingsKey = "dgvPOGView";
            this.dgvPOGView.Size = new System.Drawing.Size(520, 446);
            this.dgvPOGView.TabIndex = 18;
            this.dgvPOGView.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            this.dgvPOGView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.dgvQDView_MouseDoubleClick);
            this.dgvPOGView.RowDoubleClick += new Janus.Windows.GridEX.RowActionEventHandler(this.dgvPOGView_RowDoubleClick);
            this.dgvPOGView.FilterApplied += new System.EventHandler(this.dgvQDView_FilterApplied);
            this.dgvPOGView.GroupsChanged += new Janus.Windows.GridEX.GroupsChangedEventHandler(this.dgvQDView_GroupsChanged);
            this.dgvPOGView.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dgvQDView_KeyUp);
            // 
            // frmPOGView
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(520, 483);
            this.Controls.Add(this.dgvPOGView);
            this.Controls.Add(this.panel1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "frmPOGView";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Group List";
            this.Load += new System.EventHandler(this.Form_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvPOGView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btRefresh;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Button btSave;
        private Janus.Windows.GridEX.GridEX dgvPOGView;
    }
}