namespace QueryDesigner
{
    partial class FrmTransferOut
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
            this.radButton1 = new System.Windows.Forms.Button();
            this.radGridView1 = new Janus.Windows.GridEX.GridEX();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.radGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.radButton1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(680, 45);
            this.panel1.TabIndex = 3;
            // 
            // radButton1
            // 
            this.radButton1.Image = global::QueryDesigner.Properties.Resources._1303702678_application_vnd_ms_excel;
            this.radButton1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.radButton1.Location = new System.Drawing.Point(3, 3);
            this.radButton1.Name = "radButton1";
            this.radButton1.Size = new System.Drawing.Size(102, 37);
            this.radButton1.TabIndex = 0;
            this.radButton1.Text = "Export    ";
            this.radButton1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.radButton1.UseVisualStyleBackColor = true;
            this.radButton1.Click += new System.EventHandler(this.radButton1_Click);
            // 
            // radGridView1
            // 
            this.radGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.radGridView1.GroupByBoxVisible = false;
            this.radGridView1.Location = new System.Drawing.Point(0, 45);
            this.radGridView1.Name = "radGridView1";
            this.radGridView1.Size = new System.Drawing.Size(680, 330);
            this.radGridView1.TabIndex = 4;
            this.radGridView1.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            // 
            // FrmTransferOut
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(680, 375);
            this.Controls.Add(this.radGridView1);
            this.Controls.Add(this.panel1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Name = "FrmTransferOut";
            this.Text = "FrmTransferOut";
            this.Load += new System.EventHandler(this.FrmTransferOut_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.radGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button radButton1;
        private Janus.Windows.GridEX.GridEX radGridView1;

    }
}