namespace QueryDesigner
{
    partial class frmValidatedList
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
            Janus.Windows.GridEX.GridEXLayout ddlQD_DesignTimeLayout = new Janus.Windows.GridEX.GridEXLayout();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmValidatedList));
            Janus.Windows.GridEX.GridEXLayout ddlFld_DesignTimeLayout = new Janus.Windows.GridEX.GridEXLayout();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.ddlQD = new Janus.Windows.GridEX.EditControls.MultiColumnCombo();
            this.ddlFld = new Janus.Windows.GridEX.EditControls.MultiColumnCombo();
            this.txtMessage = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.ddlQD)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ddlFld)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(41, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "QD Code";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(242, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Field Code";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(118, 124);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(199, 124);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 5;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // ddlQD
            // 
            ddlQD_DesignTimeLayout.LayoutString = resources.GetString("ddlQD_DesignTimeLayout.LayoutString");
            this.ddlQD.DesignTimeLayout = ddlQD_DesignTimeLayout;
            this.ddlQD.DisplayMember = "QD_ID";
            this.ddlQD.Location = new System.Drawing.Point(12, 25);
            this.ddlQD.Name = "ddlQD";
            this.ddlQD.SelectedIndex = -1;
            this.ddlQD.SelectedItem = null;
            this.ddlQD.Size = new System.Drawing.Size(159, 20);
            this.ddlQD.TabIndex = 6;
            this.ddlQD.ValueMember = "QD_ID";
            this.ddlQD.ValueChanged += new System.EventHandler(this.multiColumnCombo1_ValueChanged);
            // 
            // ddlFld
            // 
            ddlFld_DesignTimeLayout.LayoutString = resources.GetString("ddlFld_DesignTimeLayout.LayoutString");
            this.ddlFld.DesignTimeLayout = ddlFld_DesignTimeLayout;
            this.ddlFld.DisplayMember = "DESCRIPTN";
            this.ddlFld.Location = new System.Drawing.Point(177, 25);
            this.ddlFld.Name = "ddlFld";
            this.ddlFld.SelectedIndex = -1;
            this.ddlFld.SelectedItem = null;
            this.ddlFld.Size = new System.Drawing.Size(212, 20);
            this.ddlFld.TabIndex = 7;
            this.ddlFld.ValueMember = "DESCRIPTN";
            // 
            // txtMessage
            // 
            this.txtMessage.Location = new System.Drawing.Point(12, 51);
            this.txtMessage.Multiline = true;
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.Size = new System.Drawing.Size(377, 67);
            this.txtMessage.TabIndex = 8;
            // 
            // frmValidatedList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(396, 159);
            this.Controls.Add(this.txtMessage);
            this.Controls.Add(this.ddlFld);
            this.Controls.Add(this.ddlQD);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "frmValidatedList";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Validated List";
            this.Load += new System.EventHandler(this.frmValidatedList_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ddlQD)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ddlFld)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private Janus.Windows.GridEX.EditControls.MultiColumnCombo ddlQD;
        private Janus.Windows.GridEX.EditControls.MultiColumnCombo ddlFld;
        private System.Windows.Forms.TextBox txtMessage;
    }
}