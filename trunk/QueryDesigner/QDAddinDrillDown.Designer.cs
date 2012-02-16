namespace QueryDesigner
{
    partial class QDAddinDrillDown
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(QDAddinDrillDown));
            this.gridEXExporter1 = new Janus.Windows.GridEX.Export.GridEXExporter(this.components);
            this.dgvResult = new Janus.Windows.GridEX.GridEX();
            this.btnExpandAll = new System.Windows.Forms.Button();
            this.btnCollapseAll = new System.Windows.Forms.Button();
            this.btnPrint = new System.Windows.Forms.Button();
            this.btPivotTable = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.twSchema1 = new System.Windows.Forms.TreeView();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lbErr = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResult)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // gridEXExporter1
            // 
            this.gridEXExporter1.GridEX = this.dgvResult;
            // 
            // dgvResult
            // 
            this.dgvResult.AllowDrop = true;
            this.dgvResult.ColumnAutoSizeMode = Janus.Windows.GridEX.ColumnAutoSizeMode.DisplayedCellsAndHeader;
            this.dgvResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvResult.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvResult.Location = new System.Drawing.Point(0, 28);
            this.dgvResult.Name = "dgvResult";
            this.dgvResult.RowHeaders = Janus.Windows.GridEX.InheritableBoolean.True;
            this.dgvResult.Size = new System.Drawing.Size(942, 514);
            this.dgvResult.TabIndex = 10;
            this.dgvResult.TotalRow = Janus.Windows.GridEX.InheritableBoolean.True;
            this.dgvResult.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2007;
            this.dgvResult.MouseLeave += new System.EventHandler(this.dgvResult_MouseLeave);
            this.dgvResult.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dgvResult_MouseUp);
            this.dgvResult.MouseMove += new System.Windows.Forms.MouseEventHandler(this.dgvResult_MouseMove);
            this.dgvResult.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dgvResult_MouseDown);
            this.dgvResult.DraggingColumn += new Janus.Windows.GridEX.ColumnActionCancelEventHandler(this.dgvResult_DraggingColumn);
            this.dgvResult.DragEnter += new System.Windows.Forms.DragEventHandler(this.dgvResult_DragEnter);
            this.dgvResult.DragDrop += new System.Windows.Forms.DragEventHandler(this.dgvResult_DragDrop);
            // 
            // btnExpandAll
            // 
            this.btnExpandAll.Location = new System.Drawing.Point(3, 1);
            this.btnExpandAll.Name = "btnExpandAll";
            this.btnExpandAll.Size = new System.Drawing.Size(75, 23);
            this.btnExpandAll.TabIndex = 5;
            this.btnExpandAll.Text = "Expand All";
            this.btnExpandAll.UseVisualStyleBackColor = true;
            this.btnExpandAll.Click += new System.EventHandler(this.btnExpandAll_Click);
            // 
            // btnCollapseAll
            // 
            this.btnCollapseAll.Location = new System.Drawing.Point(84, 1);
            this.btnCollapseAll.Name = "btnCollapseAll";
            this.btnCollapseAll.Size = new System.Drawing.Size(75, 23);
            this.btnCollapseAll.TabIndex = 6;
            this.btnCollapseAll.Text = "Collapse All";
            this.btnCollapseAll.UseVisualStyleBackColor = true;
            this.btnCollapseAll.Click += new System.EventHandler(this.btnCollapseAll_Click);
            // 
            // btnPrint
            // 
            this.btnPrint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPrint.Location = new System.Drawing.Point(700, 1);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 23);
            this.btnPrint.TabIndex = 6;
            this.btnPrint.Text = "Print";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btPivotTable
            // 
            this.btPivotTable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btPivotTable.Location = new System.Drawing.Point(781, 1);
            this.btPivotTable.Name = "btPivotTable";
            this.btPivotTable.Size = new System.Drawing.Size(75, 23);
            this.btPivotTable.TabIndex = 7;
            this.btPivotTable.Text = "Pivot Table";
            this.btPivotTable.UseVisualStyleBackColor = true;
            this.btPivotTable.Click += new System.EventHandler(this.btPivotTable_Click);
            // 
            // btCancel
            // 
            this.btCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btCancel.Location = new System.Drawing.Point(862, 1);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(75, 23);
            this.btCancel.TabIndex = 8;
            this.btCancel.Text = "Cancel";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.twSchema1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dgvResult);
            this.splitContainer1.Panel2.Controls.Add(this.panel2);
            this.splitContainer1.Panel2.Controls.Add(this.panel1);
            this.splitContainer1.Size = new System.Drawing.Size(1157, 570);
            this.splitContainer1.SplitterDistance = 211;
            this.splitContainer1.TabIndex = 10;
            // 
            // twSchema1
            // 
            this.twSchema1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.twSchema1.Location = new System.Drawing.Point(0, 0);
            this.twSchema1.Name = "twSchema1";
            this.twSchema1.Size = new System.Drawing.Size(211, 570);
            this.twSchema1.TabIndex = 0;
            this.twSchema1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.twSchema1_MouseDoubleClick);
            this.twSchema1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.twSchema1_KeyUp);
            this.twSchema1.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.twSchema1_ItemDrag);
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.lbErr);
            this.panel2.Controls.Add(this.btnPrint);
            this.panel2.Controls.Add(this.btPivotTable);
            this.panel2.Controls.Add(this.btCancel);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 542);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(942, 28);
            this.panel2.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.btnExpandAll);
            this.panel1.Controls.Add(this.btnCollapseAll);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(942, 28);
            this.panel1.TabIndex = 0;
            // 
            // lbErr
            // 
            this.lbErr.AutoSize = true;
            this.lbErr.Location = new System.Drawing.Point(3, 6);
            this.lbErr.Name = "lbErr";
            this.lbErr.Size = new System.Drawing.Size(16, 13);
            this.lbErr.TabIndex = 9;
            this.lbErr.Text = "...";
            this.lbErr.Click += new System.EventHandler(this.lbErr_Click);
            // 
            // QDAddinDrillDown
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1157, 570);
            this.Controls.Add(this.splitContainer1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(66)))), ((int)(((byte)(139)))));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "QDAddinDrillDown";
            this.Text = "Analyzer";
            this.Load += new System.EventHandler(this.QDAddin_Load);
            this.MouseUp += new System.Windows.Forms.MouseEventHandler(this.QDAddinDrillDown_MouseUp);
            ((System.ComponentModel.ISupportInitialize)(this.dgvResult)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
       
        private Janus.Windows.GridEX.Export.GridEXExporter gridEXExporter1;
        private System.Windows.Forms.Button btnCollapseAll;
        private System.Windows.Forms.Button btnExpandAll;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Button btPivotTable;
        private System.Windows.Forms.Button btnPrint;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel1;
        private Janus.Windows.GridEX.GridEX dgvResult;
        private System.Windows.Forms.TreeView twSchema1;
        private System.Windows.Forms.Label lbErr;



    }
}