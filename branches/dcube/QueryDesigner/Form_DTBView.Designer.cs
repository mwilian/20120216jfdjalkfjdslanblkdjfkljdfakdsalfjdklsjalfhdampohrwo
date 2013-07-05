namespace dCube
{
    partial class Form_DTBView
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
            Janus.Windows.GridEX.GridEXLayout dgvDTBView_Layout_0 = new Janus.Windows.GridEX.GridEXLayout();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_DTBView));
            this.btRefresh = new System.Windows.Forms.Button();
            this.btSave = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.dgvDTBView = new Janus.Windows.GridEX.GridEX();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDTBView)).BeginInit();
            this.SuspendLayout();
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
            // btSave
            // 
            this.btSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btSave.Location = new System.Drawing.Point(534, 3);
            this.btSave.Name = "btSave";
            this.btSave.Size = new System.Drawing.Size(97, 32);
            this.btSave.TabIndex = 14;
            this.btSave.Text = "OK";
            this.btSave.UseVisualStyleBackColor = true;
            this.btSave.Click += new System.EventHandler(this.btSave_Click);
            // 
            // btCancel
            // 
            this.btCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btCancel.Location = new System.Drawing.Point(637, 3);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(97, 32);
            this.btCancel.TabIndex = 15;
            this.btCancel.Text = "Cancel";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.btRefresh);
            this.panel1.Controls.Add(this.btCancel);
            this.panel1.Controls.Add(this.btSave);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 430);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(739, 37);
            this.panel1.TabIndex = 16;
            // 
            // dgvDTBView
            // 
            this.dgvDTBView.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.False;
            this.dgvDTBView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvDTBView.FilterMode = Janus.Windows.GridEX.FilterMode.Automatic;
            this.dgvDTBView.FilterRowButtonStyle = Janus.Windows.GridEX.FilterRowButtonStyle.ConditionOperatorDropDown;
            this.dgvDTBView.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            dgvDTBView_Layout_0.IsCurrentLayout = true;
            dgvDTBView_Layout_0.Key = "dgvDTBView";
            dgvDTBView_Layout_0.LayoutString = resources.GetString("dgvDTBView_Layout_0.LayoutString");
            this.dgvDTBView.Layouts.AddRange(new Janus.Windows.GridEX.GridEXLayout[] {
            dgvDTBView_Layout_0});
            this.dgvDTBView.Location = new System.Drawing.Point(0, 0);
            this.dgvDTBView.Name = "dgvDTBView";
            this.dgvDTBView.SettingsKey = "dgvDTBView";
            this.dgvDTBView.Size = new System.Drawing.Size(739, 430);
            this.dgvDTBView.TabIndex = 17;
            this.dgvDTBView.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            this.dgvDTBView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.dgvFilter_MouseDoubleClick);
            this.dgvDTBView.RowDoubleClick += new Janus.Windows.GridEX.RowActionEventHandler(this.dgvDTBView_RowDoubleClick);
            this.dgvDTBView.ColumnMoved += new Janus.Windows.GridEX.ColumnActionEventHandler(this.dgvDTBView_ColumnMoved);
            this.dgvDTBView.FilterApplied += new System.EventHandler(this.dgvDTBView_FilterApplied);
            this.dgvDTBView.GroupsChanged += new Janus.Windows.GridEX.GroupsChangedEventHandler(this.dgvDTBView_GroupsChanged);
            this.dgvDTBView.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dgvFilter_KeyUp);
            // 
            // Form_DTBView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(739, 467);
            this.Controls.Add(this.dgvDTBView);
            this.Controls.Add(this.panel1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "Form_DTBView";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Database View";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvDTBView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btRefresh;
        private System.Windows.Forms.Button btSave;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Panel panel1;
        private Janus.Windows.GridEX.GridEX dgvDTBView;

    }
}