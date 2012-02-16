namespace QueryDesigner
{
    partial class frmChangeDB
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
            this.bt_database = new System.Windows.Forms.PictureBox();
            this.txt_database = new System.Windows.Forms.Label();
            this.lblDatabase = new System.Windows.Forms.Label();
            this.txtdatabase = new System.Windows.Forms.TextBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.bt_database)).BeginInit();
            this.SuspendLayout();
            // 
            // bt_database
            // 
            this.bt_database.BackColor = System.Drawing.Color.Transparent;
            this.bt_database.Image = global::QueryDesigner.Properties.Resources._1303882176_search_16;
            this.bt_database.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.bt_database.Location = new System.Drawing.Point(130, 20);
            this.bt_database.Name = "bt_database";
            this.bt_database.Size = new System.Drawing.Size(16, 16);
            this.bt_database.TabIndex = 49;
            this.bt_database.TabStop = false;
            this.bt_database.Click += new System.EventHandler(this.bt_database_Click);
            // 
            // txt_database
            // 
            this.txt_database.AutoSize = true;
            this.txt_database.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.txt_database.Location = new System.Drawing.Point(163, 21);
            this.txt_database.Name = "txt_database";
            this.txt_database.Size = new System.Drawing.Size(25, 13);
            this.txt_database.TabIndex = 48;
            this.txt_database.Text = "___";
            // 
            // lblDatabase
            // 
            this.lblDatabase.AutoSize = true;
            this.lblDatabase.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.lblDatabase.Location = new System.Drawing.Point(12, 21);
            this.lblDatabase.Name = "lblDatabase";
            this.lblDatabase.Size = new System.Drawing.Size(53, 13);
            this.lblDatabase.TabIndex = 47;
            this.lblDatabase.Text = "Database";
            // 
            // txtdatabase
            // 
            this.txtdatabase.Location = new System.Drawing.Point(71, 18);
            this.txtdatabase.Name = "txtdatabase";
            this.txtdatabase.Size = new System.Drawing.Size(53, 20);
            this.txtdatabase.TabIndex = 46;
            this.txtdatabase.Validated += new System.EventHandler(this.txtdatabase_Validated);
            this.txtdatabase.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtdatabase_KeyUp);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(71, 49);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(79, 17);
            this.checkBox1.TabIndex = 51;
            this.checkBox1.Text = "Set Default";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(39, 79);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(100, 23);
            this.btnOK.TabIndex = 52;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(177, 79);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(100, 23);
            this.btnCancel.TabIndex = 53;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // frmChangeDB
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(316, 114);
            this.ControlBox = false;
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.bt_database);
            this.Controls.Add(this.txt_database);
            this.Controls.Add(this.lblDatabase);
            this.Controls.Add(this.txtdatabase);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "frmChangeDB";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Change DB";
            this.Load += new System.EventHandler(this.frmChangeDB_Load);
            ((System.ComponentModel.ISupportInitialize)(this.bt_database)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox bt_database;
        private System.Windows.Forms.Label txt_database;
        private System.Windows.Forms.Label lblDatabase;
        private System.Windows.Forms.TextBox txtdatabase;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
    }
}