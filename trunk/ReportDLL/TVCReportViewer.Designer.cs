namespace ReportDLL
{
    partial class TVCReportViewer
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TVCReportViewer));
            this.panel1 = new System.Windows.Forms.Panel();
            this.flexCelPreview1 = new FlexCel.Winforms.FlexCelPreview();
            this.thumbs = new FlexCel.Winforms.FlexCelPreview();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.panelLeft = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.cbAntiAlias = new System.Windows.Forms.ComboBox();
            this.btnZoomIn = new System.Windows.Forms.Button();
            this.btnZoomOut = new System.Windows.Forms.Button();
            this.btnLast = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnFirst = new System.Windows.Forms.Button();
            this.btnPrev = new System.Windows.Forms.Button();
            this.edZoom = new System.Windows.Forms.TextBox();
            this.edPage = new System.Windows.Forms.TextBox();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.flexCelImgExport1 = new FlexCel.Render.FlexCelImgExport();
            this.PdfSaveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.printDialog1 = new System.Windows.Forms.PrintDialog();
            this.btnPrintReport = new System.Windows.Forms.Button();
            this.btnRecalc = new System.Windows.Forms.Button();
            this.btnPdf = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.panelLeft.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.flexCelPreview1);
            this.panel1.Controls.Add(this.splitter1);
            this.panel1.Controls.Add(this.panelLeft);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 34);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(897, 368);
            this.panel1.TabIndex = 12;
            // 
            // flexCelPreview1
            // 
            this.flexCelPreview1.AutoScroll = true;
            this.flexCelPreview1.AutoScrollMinSize = new System.Drawing.Size(40, 10);
            this.flexCelPreview1.BackColor = System.Drawing.Color.Gray;
            this.flexCelPreview1.CacheSize = 64;
            this.flexCelPreview1.CenteredPreview = false;
            this.flexCelPreview1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flexCelPreview1.Document = this.flexCelImgExport1;
            this.flexCelPreview1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            this.flexCelPreview1.Location = new System.Drawing.Point(144, 0);
            this.flexCelPreview1.Name = "flexCelPreview1";
            this.flexCelPreview1.PageXSeparation = 20;
            this.flexCelPreview1.PageYSeparation = 10;
            this.flexCelPreview1.Size = new System.Drawing.Size(751, 366);
            this.flexCelPreview1.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            this.flexCelPreview1.StartPage = 1;
            this.flexCelPreview1.TabIndex = 2;
            this.flexCelPreview1.ThumbnailLarge = null;
            this.flexCelPreview1.ThumbnailSmall = this.thumbs;
            this.flexCelPreview1.Zoom = 1F;
            this.flexCelPreview1.StartPageChanged += new System.EventHandler(this.flexCelPreview1_StartPageChanged);
            this.flexCelPreview1.ZoomChanged += new System.EventHandler(this.flexCelPreview1_ZoomChanged);
            // 
            // thumbs
            // 
            this.thumbs.AutoScroll = true;
            this.thumbs.AutoScrollMinSize = new System.Drawing.Size(20, 10);
            this.thumbs.BackColor = System.Drawing.Color.Gray;
            this.thumbs.CacheSize = 64;
            this.thumbs.CenteredPreview = false;
            this.thumbs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.thumbs.Document = this.flexCelImgExport1;
            this.thumbs.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Default;
            this.thumbs.Location = new System.Drawing.Point(0, 16);
            this.thumbs.Name = "thumbs";
            this.thumbs.PageXSeparation = 10;
            this.thumbs.PageYSeparation = 10;
            this.thumbs.Size = new System.Drawing.Size(136, 350);
            this.thumbs.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
            this.thumbs.StartPage = 1;
            this.thumbs.TabIndex = 3;
            this.thumbs.ThumbnailLarge = this.flexCelPreview1;
            this.thumbs.ThumbnailSmall = null;
            this.thumbs.Zoom = 0.1F;
            // 
            // splitter1
            // 
            this.splitter1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.splitter1.Location = new System.Drawing.Point(136, 0);
            this.splitter1.MinSize = 0;
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(8, 366);
            this.splitter1.TabIndex = 11;
            this.splitter1.TabStop = false;
            // 
            // panelLeft
            // 
            this.panelLeft.Controls.Add(this.thumbs);
            this.panelLeft.Controls.Add(this.label2);
            this.panelLeft.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelLeft.Location = new System.Drawing.Point(0, 0);
            this.panelLeft.Name = "panelLeft";
            this.panelLeft.Size = new System.Drawing.Size(136, 366);
            this.panelLeft.TabIndex = 9;
            // 
            // label2
            // 
            this.label2.Dock = System.Windows.Forms.DockStyle.Top;
            this.label2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label2.Location = new System.Drawing.Point(0, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(136, 16);
            this.label2.TabIndex = 13;
            this.label2.Text = "Thumbs";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnPrintReport);
            this.panel2.Controls.Add(this.cbAntiAlias);
            this.panel2.Controls.Add(this.btnRecalc);
            this.panel2.Controls.Add(this.btnPdf);
            this.panel2.Controls.Add(this.btnZoomIn);
            this.panel2.Controls.Add(this.btnZoomOut);
            this.panel2.Controls.Add(this.btnLast);
            this.panel2.Controls.Add(this.btnNext);
            this.panel2.Controls.Add(this.btnFirst);
            this.panel2.Controls.Add(this.btnPrev);
            this.panel2.Controls.Add(this.edZoom);
            this.panel2.Controls.Add(this.edPage);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(897, 34);
            this.panel2.TabIndex = 11;
            // 
            // cbAntiAlias
            // 
            this.cbAntiAlias.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbAntiAlias.Items.AddRange(new object[] {
            "Antialiased",
            "High Quality",
            "High Speed"});
            this.cbAntiAlias.Location = new System.Drawing.Point(603, 6);
            this.cbAntiAlias.Name = "cbAntiAlias";
            this.cbAntiAlias.Size = new System.Drawing.Size(121, 21);
            this.cbAntiAlias.TabIndex = 18;
            this.cbAntiAlias.SelectedIndexChanged += new System.EventHandler(this.cbAntiAlias_SelectedIndexChanged);
            // 
            // btnZoomIn
            // 
            this.btnZoomIn.BackColor = System.Drawing.SystemColors.Control;
            this.btnZoomIn.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnZoomIn.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnZoomIn.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnZoomIn.ImageIndex = 3;
            this.btnZoomIn.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnZoomIn.Location = new System.Drawing.Point(571, 6);
            this.btnZoomIn.Name = "btnZoomIn";
            this.btnZoomIn.Size = new System.Drawing.Size(20, 20);
            this.btnZoomIn.TabIndex = 14;
            this.btnZoomIn.Text = "+";
            this.btnZoomIn.UseVisualStyleBackColor = false;
            this.btnZoomIn.Click += new System.EventHandler(this.btnZoomIn_Click);
            // 
            // btnZoomOut
            // 
            this.btnZoomOut.BackColor = System.Drawing.SystemColors.Control;
            this.btnZoomOut.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnZoomOut.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnZoomOut.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnZoomOut.ImageIndex = 3;
            this.btnZoomOut.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnZoomOut.Location = new System.Drawing.Point(485, 6);
            this.btnZoomOut.Name = "btnZoomOut";
            this.btnZoomOut.Size = new System.Drawing.Size(20, 20);
            this.btnZoomOut.TabIndex = 13;
            this.btnZoomOut.Text = "-";
            this.btnZoomOut.UseVisualStyleBackColor = false;
            this.btnZoomOut.Click += new System.EventHandler(this.btnZoomOut_Click);
            // 
            // btnLast
            // 
            this.btnLast.BackColor = System.Drawing.SystemColors.Control;
            this.btnLast.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnLast.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnLast.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnLast.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnLast.Location = new System.Drawing.Point(446, 6);
            this.btnLast.Name = "btnLast";
            this.btnLast.Size = new System.Drawing.Size(20, 20);
            this.btnLast.TabIndex = 12;
            this.btnLast.Text = ">>";
            this.btnLast.UseVisualStyleBackColor = false;
            this.btnLast.Click += new System.EventHandler(this.btnLast_Click);
            // 
            // btnNext
            // 
            this.btnNext.BackColor = System.Drawing.SystemColors.Control;
            this.btnNext.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnNext.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnNext.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnNext.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnNext.Location = new System.Drawing.Point(427, 6);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(20, 20);
            this.btnNext.TabIndex = 11;
            this.btnNext.Text = ">";
            this.btnNext.UseVisualStyleBackColor = false;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnFirst
            // 
            this.btnFirst.BackColor = System.Drawing.SystemColors.Control;
            this.btnFirst.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnFirst.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnFirst.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnFirst.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnFirst.Location = new System.Drawing.Point(276, 6);
            this.btnFirst.Name = "btnFirst";
            this.btnFirst.Size = new System.Drawing.Size(20, 20);
            this.btnFirst.TabIndex = 10;
            this.btnFirst.Text = "<<";
            this.btnFirst.UseVisualStyleBackColor = false;
            this.btnFirst.Click += new System.EventHandler(this.btnFirst_Click);
            // 
            // btnPrev
            // 
            this.btnPrev.BackColor = System.Drawing.SystemColors.Control;
            this.btnPrev.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnPrev.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnPrev.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPrev.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnPrev.Location = new System.Drawing.Point(302, 6);
            this.btnPrev.Name = "btnPrev";
            this.btnPrev.Size = new System.Drawing.Size(21, 20);
            this.btnPrev.TabIndex = 9;
            this.btnPrev.Text = "<";
            this.btnPrev.UseVisualStyleBackColor = false;
            this.btnPrev.Click += new System.EventHandler(this.btnPrev_Click);
            // 
            // edZoom
            // 
            this.edZoom.Location = new System.Drawing.Point(507, 6);
            this.edZoom.Name = "edZoom";
            this.edZoom.Size = new System.Drawing.Size(64, 20);
            this.edZoom.TabIndex = 8;
            this.edZoom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.edZoom.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.edZoom_KeyPress);
            this.edZoom.Enter += new System.EventHandler(this.edZoom_Enter);
            // 
            // edPage
            // 
            this.edPage.Location = new System.Drawing.Point(323, 6);
            this.edPage.Name = "edPage";
            this.edPage.Size = new System.Drawing.Size(104, 20);
            this.edPage.TabIndex = 7;
            this.edPage.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.edPage.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.edPage_KeyPress);
            this.edPage.Enter += new System.EventHandler(this.edPage_Leave);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "Excel Files|*.xls";
            // 
            // flexCelImgExport1
            // 
            this.flexCelImgExport1.AllVisibleSheets = false;
            this.flexCelImgExport1.PageSize = null;
            this.flexCelImgExport1.ResetPageNumberOnEachSheet = false;
            this.flexCelImgExport1.Resolution = 96F;
            this.flexCelImgExport1.Workbook = null;
            // 
            // PdfSaveFileDialog
            // 
            this.PdfSaveFileDialog.DefaultExt = "pdf";
            this.PdfSaveFileDialog.Filter = "Pdf Files|*.pdf";
            this.PdfSaveFileDialog.Title = "Select the file to export to:";
            // 
            // openFileDialog
            // 
            this.openFileDialog.DefaultExt = "xls";
            this.openFileDialog.Filter = "Excel Files|*.xls|All files|*.*";
            this.openFileDialog.Title = "Select a file to preview";
            // 
            // printDialog1
            // 
            this.printDialog1.AllowSomePages = true;
            this.printDialog1.UseEXDialog = true;
            // 
            // btnPrintReport
            // 
            this.btnPrintReport.Image = global::ReportDLL.Properties.Resources._1303701366_print_32;
            this.btnPrintReport.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPrintReport.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnPrintReport.Location = new System.Drawing.Point(5, 2);
            this.btnPrintReport.Name = "btnPrintReport";
            this.btnPrintReport.Size = new System.Drawing.Size(83, 30);
            this.btnPrintReport.TabIndex = 19;
            this.btnPrintReport.Text = "Print  ";
            this.btnPrintReport.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnPrintReport.UseVisualStyleBackColor = true;
            this.btnPrintReport.Click += new System.EventHandler(this.btnPrintReport_Click);
            // 
            // btnRecalc
            // 
            this.btnRecalc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRecalc.BackColor = System.Drawing.SystemColors.Control;
            this.btnRecalc.Image = ((System.Drawing.Image)(resources.GetObject("btnRecalc.Image")));
            this.btnRecalc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRecalc.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnRecalc.Location = new System.Drawing.Point(729, 2);
            this.btnRecalc.Name = "btnRecalc";
            this.btnRecalc.Size = new System.Drawing.Size(72, 30);
            this.btnRecalc.TabIndex = 17;
            this.btnRecalc.Text = "Recalc";
            this.btnRecalc.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnRecalc.UseVisualStyleBackColor = false;
            this.btnRecalc.Click += new System.EventHandler(this.btnRecalc_Click);
            // 
            // btnPdf
            // 
            this.btnPdf.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPdf.BackColor = System.Drawing.SystemColors.Control;
            this.btnPdf.Image = ((System.Drawing.Image)(resources.GetObject("btnPdf.Image")));
            this.btnPdf.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPdf.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnPdf.Location = new System.Drawing.Point(809, 2);
            this.btnPdf.Name = "btnPdf";
            this.btnPdf.Size = new System.Drawing.Size(56, 30);
            this.btnPdf.TabIndex = 16;
            this.btnPdf.Text = "Pdf";
            this.btnPdf.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnPdf.UseVisualStyleBackColor = false;
            this.btnPdf.Click += new System.EventHandler(this.btnPdf_Click);
            // 
            // TVCReportViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Name = "TVCReportViewer";
            this.Size = new System.Drawing.Size(897, 402);
            this.panel1.ResumeLayout(false);
            this.panelLeft.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private FlexCel.Winforms.FlexCelPreview flexCelPreview1;
        private FlexCel.Winforms.FlexCelPreview thumbs;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.Panel panelLeft;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btnPrintReport;
        private System.Windows.Forms.ComboBox cbAntiAlias;
        private System.Windows.Forms.Button btnRecalc;
        private System.Windows.Forms.Button btnPdf;
        private System.Windows.Forms.Button btnZoomIn;
        private System.Windows.Forms.Button btnZoomOut;
        private System.Windows.Forms.Button btnLast;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnFirst;
        private System.Windows.Forms.Button btnPrev;
        private System.Windows.Forms.TextBox edZoom;
        private System.Windows.Forms.TextBox edPage;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private FlexCel.Render.FlexCelImgExport flexCelImgExport1;
        private System.Windows.Forms.SaveFileDialog PdfSaveFileDialog;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.PrintDialog printDialog1;




    }
}
