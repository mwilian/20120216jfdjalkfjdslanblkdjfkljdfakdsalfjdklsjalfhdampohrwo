namespace dCube
{
    partial class Form_TableView
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_TableView));
            Janus.Windows.GridEX.GridEXLayout dgvTableView_Layout_0 = new Janus.Windows.GridEX.GridEXLayout();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btRefresh = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.btSave = new System.Windows.Forms.Button();
            this.dgvTableView = new Janus.Windows.GridEX.GridEX();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTableView)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.btRefresh);
            this.panel1.Controls.Add(this.btCancel);
            this.panel1.Controls.Add(this.btSave);
            resources.ApplyResources(this.panel1, "panel1");
            this.panel1.Name = "panel1";
            // 
            // btRefresh
            // 
            resources.ApplyResources(this.btRefresh, "btRefresh");
            this.btRefresh.Name = "btRefresh";
            this.btRefresh.UseVisualStyleBackColor = true;
            this.btRefresh.Click += new System.EventHandler(this.btnReresh_Click);
            // 
            // btCancel
            // 
            resources.ApplyResources(this.btCancel, "btCancel");
            this.btCancel.Name = "btCancel";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // btSave
            // 
            resources.ApplyResources(this.btSave, "btSave");
            this.btSave.Name = "btSave";
            this.btSave.UseVisualStyleBackColor = true;
            this.btSave.Click += new System.EventHandler(this.btSave_Click);
            // 
            // dgvTableView
            // 
            this.dgvTableView.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.False;
            resources.ApplyResources(this.dgvTableView, "dgvTableView");
            this.dgvTableView.FilterMode = Janus.Windows.GridEX.FilterMode.Automatic;
            this.dgvTableView.FilterRowButtonStyle = Janus.Windows.GridEX.FilterRowButtonStyle.ConditionOperatorDropDown;
            dgvTableView_Layout_0.IsCurrentLayout = true;
            dgvTableView_Layout_0.Key = "dgvTableView";
            resources.ApplyResources(dgvTableView_Layout_0, "dgvTableView_Layout_0");
            this.dgvTableView.Layouts.AddRange(new Janus.Windows.GridEX.GridEXLayout[] {
            dgvTableView_Layout_0});
            this.dgvTableView.Name = "dgvTableView";
            this.dgvTableView.SettingsKey = "dgvTableView";
            this.dgvTableView.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            this.dgvTableView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.dgvTableView_MouseDoubleClick);
            this.dgvTableView.RowDoubleClick += new Janus.Windows.GridEX.RowActionEventHandler(this.dgvTableView_RowDoubleClick);
            this.dgvTableView.FilterApplied += new System.EventHandler(this.dgvTableView_FilterApplied);
            this.dgvTableView.GroupsChanged += new Janus.Windows.GridEX.GroupsChangedEventHandler(this.dgvTableView_GroupsChanged);
            this.dgvTableView.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dgvTableView_KeyUp);
            // 
            // Form_TableView
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dgvTableView);
            this.Controls.Add(this.panel1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "Form_TableView";
            this.Load += new System.EventHandler(this.Form_TableView_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTableView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btRefresh;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Button btSave;
        private Janus.Windows.GridEX.GridEX dgvTableView;



    }
}