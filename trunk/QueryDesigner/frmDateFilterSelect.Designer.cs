namespace dCube
{
    partial class frmDateFilterSelect
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtFilterTo = new System.Windows.Forms.DateTimePicker();
            this.txtFilterFrom = new System.Windows.Forms.DateTimePicker();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txtFilterTo);
            this.panel1.Controls.Add(this.txtFilterFrom);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.btnOK);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel1.Location = new System.Drawing.Point(470, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(243, 390);
            this.panel1.TabIndex = 3;
            // 
            // txtFilterTo
            // 
            this.txtFilterTo.CustomFormat = "dd/MM/yyyy";
            this.txtFilterTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.txtFilterTo.Location = new System.Drawing.Point(12, 79);
            this.txtFilterTo.MinDate = new System.DateTime(1900, 1, 1, 0, 0, 0, 0);
            this.txtFilterTo.Name = "txtFilterTo";
            this.txtFilterTo.Size = new System.Drawing.Size(219, 20);
            this.txtFilterTo.TabIndex = 7;
            this.txtFilterTo.Value = new System.DateTime(2011, 6, 28, 23, 54, 57, 198);
            // 
            // txtFilterFrom
            // 
            this.txtFilterFrom.CustomFormat = "dd/MM/yyyy";
            this.txtFilterFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.txtFilterFrom.Location = new System.Drawing.Point(12, 28);
            this.txtFilterFrom.MinDate = new System.DateTime(1900, 1, 1, 0, 0, 0, 0);
            this.txtFilterFrom.Name = "txtFilterFrom";
            this.txtFilterFrom.Size = new System.Drawing.Size(219, 20);
            this.txtFilterFrom.TabIndex = 6;
            this.txtFilterFrom.Value = new System.DateTime(2011, 6, 28, 23, 54, 57, 198);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(131, 118);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(100, 23);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(12, 118);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(100, 23);
            this.btnOK.TabIndex = 4;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(9, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Filter To";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(9, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Filter From";
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.CalendarDimensions = new System.Drawing.Size(2, 2);
            this.monthCalendar1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.monthCalendar1.Location = new System.Drawing.Point(0, 0);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.ShowToday = false;
            this.monthCalendar1.TabIndex = 5;
            this.monthCalendar1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.monthCalendar1_MouseDown);
            // 
            // frmDateFilterSelect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(713, 390);
            this.Controls.Add(this.monthCalendar1);
            this.Controls.Add(this.panel1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "frmDateFilterSelect";
            this.Text = "Filter Selection";
            this.Load += new System.EventHandler(this.frmDateFilterSelect_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MonthCalendar monthCalendar1;
        private System.Windows.Forms.DateTimePicker txtFilterFrom;
        private System.Windows.Forms.DateTimePicker txtFilterTo;

    }
}

