using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Render;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
namespace QueryDesigner
{
	public partial class frmFlexPrintPreview : System.Windows.Forms.Form
	{
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button print;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox edFileName;
        private System.Windows.Forms.CheckBox chFormulaText;
        private System.Windows.Forms.CheckBox chAntiAlias;
        private System.Windows.Forms.Button setup;
        private System.Windows.Forms.CheckBox chGridLines;
        private System.Windows.Forms.TextBox edHeader;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox edFooter;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox edHPages;
        private System.Windows.Forms.TextBox edVPages;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox chPrintLeft;
        private System.Windows.Forms.CheckBox chFitIn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox edZoom;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox edl;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox edt;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox edr;
        private System.Windows.Forms.Label labelb;
        private System.Windows.Forms.TextBox edb;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox edh;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox edf;
        private System.Windows.Forms.CheckBox Landscape;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox edTop;
        private System.Windows.Forms.TextBox edLeft;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox edRight;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox edBottom;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.ComboBox cbSheet;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.CheckBox cbConfidential;
        private System.Windows.Forms.Button export;
        private System.Windows.Forms.SaveFileDialog exportImageDialog;
		private System.Windows.Forms.CheckBox chHeadings;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.ComboBox cbInterpolation;
		private System.Windows.Forms.SaveFileDialog exportTiffDialog;
        private System.Windows.Forms.ContextMenu exportImagesMenu;
        private System.Windows.Forms.MenuItem ExportUsingFlexCelImgExport;
        private System.Windows.Forms.MenuItem ExportMultiPageTiff;
        private System.Windows.Forms.MenuItem ExportUsingPrintController;
        private System.Windows.Forms.MenuItem ImgBlackAndWhite;
        private System.Windows.Forms.MenuItem Img256Colors;
        private System.Windows.Forms.MenuItem ImgTrueColor;
        private System.Windows.Forms.MenuItem TiffBlackAndWhite;
        private System.Windows.Forms.MenuItem TiffFax;
        private System.Windows.Forms.MenuItem Tiff256Colors;
        private System.Windows.Forms.MenuItem TiffTrueColor;
        private System.Windows.Forms.MenuItem menuItem1;
        private System.Windows.Forms.MenuItem ExportUsingSaveAsImage;
        private System.Windows.Forms.MenuItem ImgBlackAndWhite2;
        private System.Windows.Forms.MenuItem Img256Colors2;
        private System.Windows.Forms.MenuItem ImgTrueColor2;
        private System.Windows.Forms.MenuItem menuItem6;
		private System.Windows.Forms.CheckBox cbAllSheets;
		private System.Windows.Forms.Panel panel4;
		private System.Windows.Forms.CheckBox cbResetPageNumber;
        private System.ComponentModel.IContainer components;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmFlexPrintPreview));
            this.panel2 = new System.Windows.Forms.Panel();
            this.preview = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.export = new System.Windows.Forms.Button();
            this.setup = new System.Windows.Forms.Button();
            this.print = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.flexCelPrintDocument1 = new FlexCel.Render.FlexCelPrintDocument();
            this.panel1 = new System.Windows.Forms.Panel();
            this.cbResetPageNumber = new System.Windows.Forms.CheckBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.cbAllSheets = new System.Windows.Forms.CheckBox();
            this.label19 = new System.Windows.Forms.Label();
            this.cbInterpolation = new System.Windows.Forms.ComboBox();
            this.chHeadings = new System.Windows.Forms.CheckBox();
            this.cbConfidential = new System.Windows.Forms.CheckBox();
            this.label18 = new System.Windows.Forms.Label();
            this.cbSheet = new System.Windows.Forms.ComboBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.edBottom = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.edRight = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.edLeft = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.edTop = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.Landscape = new System.Windows.Forms.CheckBox();
            this.label11 = new System.Windows.Forms.Label();
            this.edf = new System.Windows.Forms.TextBox();
            this.labelb = new System.Windows.Forms.Label();
            this.edb = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.edr = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.edt = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.edl = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.edZoom = new System.Windows.Forms.TextBox();
            this.chFitIn = new System.Windows.Forms.CheckBox();
            this.chPrintLeft = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.edVPages = new System.Windows.Forms.TextBox();
            this.edHPages = new System.Windows.Forms.TextBox();
            this.edFooter = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.edHeader = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.edFileName = new System.Windows.Forms.TextBox();
            this.chFormulaText = new System.Windows.Forms.CheckBox();
            this.chGridLines = new System.Windows.Forms.CheckBox();
            this.chAntiAlias = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.edh = new System.Windows.Forms.TextBox();
            this.exportImageDialog = new System.Windows.Forms.SaveFileDialog();
            this.exportTiffDialog = new System.Windows.Forms.SaveFileDialog();
            this.exportImagesMenu = new System.Windows.Forms.ContextMenu();
            this.ExportUsingPrintController = new System.Windows.Forms.MenuItem();
            this.menuItem6 = new System.Windows.Forms.MenuItem();
            this.ExportUsingFlexCelImgExport = new System.Windows.Forms.MenuItem();
            this.ImgBlackAndWhite = new System.Windows.Forms.MenuItem();
            this.Img256Colors = new System.Windows.Forms.MenuItem();
            this.ImgTrueColor = new System.Windows.Forms.MenuItem();
            this.ExportUsingSaveAsImage = new System.Windows.Forms.MenuItem();
            this.ImgBlackAndWhite2 = new System.Windows.Forms.MenuItem();
            this.Img256Colors2 = new System.Windows.Forms.MenuItem();
            this.ImgTrueColor2 = new System.Windows.Forms.MenuItem();
            this.ExportMultiPageTiff = new System.Windows.Forms.MenuItem();
            this.TiffFax = new System.Windows.Forms.MenuItem();
            this.TiffBlackAndWhite = new System.Windows.Forms.MenuItem();
            this.Tiff256Colors = new System.Windows.Forms.MenuItem();
            this.TiffTrueColor = new System.Windows.Forms.MenuItem();
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.printDialog1 = new System.Windows.Forms.PrintDialog();
            this.printPreviewDialog1 = new System.Windows.Forms.PrintPreviewDialog();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.preview);
            this.panel2.Controls.Add(this.export);
            this.panel2.Controls.Add(this.setup);
            this.panel2.Controls.Add(this.print);
            this.panel2.Controls.Add(this.button2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(768, 34);
            this.panel2.TabIndex = 2;
            // 
            // preview
            // 
            this.preview.BackColor = System.Drawing.SystemColors.Control;
            this.preview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.preview.ImageIndex = 3;
            this.preview.ImageList = this.imageList1;
            this.preview.Location = new System.Drawing.Point(286, 2);
            this.preview.Name = "preview";
            this.preview.Size = new System.Drawing.Size(80, 30);
            this.preview.TabIndex = 6;
            this.preview.Text = "Preview";
            this.preview.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.preview.UseVisualStyleBackColor = false;
            this.preview.Click += new System.EventHandler(this.preview_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Magenta;
            this.imageList1.Images.SetKeyName(0, "");
            this.imageList1.Images.SetKeyName(1, "");
            this.imageList1.Images.SetKeyName(2, "");
            this.imageList1.Images.SetKeyName(3, "");
            this.imageList1.Images.SetKeyName(4, "");
            this.imageList1.Images.SetKeyName(5, "");
            this.imageList1.Images.SetKeyName(6, "");
            // 
            // export
            // 
            this.export.BackColor = System.Drawing.SystemColors.Control;
            this.export.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.export.ImageIndex = 6;
            this.export.ImageList = this.imageList1;
            this.export.Location = new System.Drawing.Point(160, 2);
            this.export.Name = "export";
            this.export.Size = new System.Drawing.Size(120, 30);
            this.export.TabIndex = 5;
            this.export.Text = "Export as Images";
            this.export.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.export.UseVisualStyleBackColor = false;
            this.export.Click += new System.EventHandler(this.export_Click);
            // 
            // setup
            // 
            this.setup.BackColor = System.Drawing.SystemColors.Control;
            this.setup.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.setup.ImageIndex = 5;
            this.setup.ImageList = this.imageList1;
            this.setup.Location = new System.Drawing.Point(3, 2);
            this.setup.Name = "setup";
            this.setup.Size = new System.Drawing.Size(80, 30);
            this.setup.TabIndex = 4;
            this.setup.Text = "Setup";
            this.setup.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.setup.UseVisualStyleBackColor = false;
            this.setup.Click += new System.EventHandler(this.setup_Click);
            // 
            // print
            // 
            this.print.BackColor = System.Drawing.SystemColors.Control;
            this.print.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.print.ImageIndex = 4;
            this.print.ImageList = this.imageList1;
            this.print.Location = new System.Drawing.Point(89, 2);
            this.print.Name = "print";
            this.print.Size = new System.Drawing.Size(64, 30);
            this.print.TabIndex = 3;
            this.print.Text = "Print";
            this.print.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.print.UseVisualStyleBackColor = false;
            this.print.Click += new System.EventHandler(this.print_Click);
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.BackColor = System.Drawing.SystemColors.Control;
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.ImageIndex = 2;
            this.button2.ImageList = this.imageList1;
            this.button2.Location = new System.Drawing.Point(708, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(56, 26);
            this.button2.TabIndex = 2;
            this.button2.Text = "Exit";
            this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "xls";
            this.openFileDialog1.Filter = "Excel Files|*.xls";
            this.openFileDialog1.Title = "Open an Excel File";
            // 
            // flexCelPrintDocument1
            // 
            this.flexCelPrintDocument1.AllVisibleSheets = false;
            this.flexCelPrintDocument1.ResetPageNumberOnEachSheet = false;
            this.flexCelPrintDocument1.Workbook = null;
            this.flexCelPrintDocument1.GetPrinterHardMargins += new FlexCel.Render.PrintHardMarginsEventHandler(this.flexCelPrintDocument1_GetPrinterHardMargins);
            this.flexCelPrintDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.flexCelPrintDocument1_PrintPage);
            this.flexCelPrintDocument1.BeforePrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.flexCelPrintDocument1_BeforePrintPage);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.cbResetPageNumber);
            this.panel1.Controls.Add(this.panel4);
            this.panel1.Controls.Add(this.cbAllSheets);
            this.panel1.Controls.Add(this.label19);
            this.panel1.Controls.Add(this.cbInterpolation);
            this.panel1.Controls.Add(this.chHeadings);
            this.panel1.Controls.Add(this.cbConfidential);
            this.panel1.Controls.Add(this.label18);
            this.panel1.Controls.Add(this.cbSheet);
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.Landscape);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.edf);
            this.panel1.Controls.Add(this.labelb);
            this.panel1.Controls.Add(this.edb);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.edr);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.edt);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.edl);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.edZoom);
            this.panel1.Controls.Add(this.chFitIn);
            this.panel1.Controls.Add(this.chPrintLeft);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.edVPages);
            this.panel1.Controls.Add(this.edHPages);
            this.panel1.Controls.Add(this.edFooter);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.edHeader);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.edFileName);
            this.panel1.Controls.Add(this.chFormulaText);
            this.panel1.Controls.Add(this.chGridLines);
            this.panel1.Controls.Add(this.chAntiAlias);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.edh);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 34);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(768, 483);
            this.panel1.TabIndex = 3;
            // 
            // cbResetPageNumber
            // 
            this.cbResetPageNumber.Enabled = false;
            this.cbResetPageNumber.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbResetPageNumber.Location = new System.Drawing.Point(528, 48);
            this.cbResetPageNumber.Name = "cbResetPageNumber";
            this.cbResetPageNumber.Size = new System.Drawing.Size(216, 16);
            this.cbResetPageNumber.TabIndex = 39;
            this.cbResetPageNumber.Text = "Reset Page number on each sheet.";
            // 
            // panel4
            // 
            this.panel4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel4.Location = new System.Drawing.Point(16, 72);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(736, 3);
            this.panel4.TabIndex = 38;
            // 
            // cbAllSheets
            // 
            this.cbAllSheets.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbAllSheets.Location = new System.Drawing.Point(32, 48);
            this.cbAllSheets.Name = "cbAllSheets";
            this.cbAllSheets.Size = new System.Drawing.Size(104, 16);
            this.cbAllSheets.TabIndex = 37;
            this.cbAllSheets.Text = "All Sheets";
            this.cbAllSheets.CheckedChanged += new System.EventHandler(this.cbAllSheets_CheckedChanged);
            // 
            // label19
            // 
            this.label19.Location = new System.Drawing.Point(392, 80);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(160, 40);
            this.label19.TabIndex = 36;
            this.label19.Text = "Interpolation mode for images: Sometimes a lower mode might give crisper results." +
                "";
            // 
            // cbInterpolation
            // 
            this.cbInterpolation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbInterpolation.Items.AddRange(new object[] {
            "Bicubic",
            "Bilinear",
            "Default",
            "High",
            "HighQualityBicubic",
            "HighQualityBilinear ",
            "Low",
            "NearestNeighbor"});
            this.cbInterpolation.Location = new System.Drawing.Point(560, 88);
            this.cbInterpolation.Name = "cbInterpolation";
            this.cbInterpolation.Size = new System.Drawing.Size(152, 21);
            this.cbInterpolation.TabIndex = 35;
            // 
            // chHeadings
            // 
            this.chHeadings.Location = new System.Drawing.Point(176, 136);
            this.chHeadings.Name = "chHeadings";
            this.chHeadings.Size = new System.Drawing.Size(128, 24);
            this.chHeadings.TabIndex = 34;
            this.chHeadings.Text = "Print Headings";
            // 
            // cbConfidential
            // 
            this.cbConfidential.Location = new System.Drawing.Point(56, 112);
            this.cbConfidential.Name = "cbConfidential";
            this.cbConfidential.Size = new System.Drawing.Size(232, 16);
            this.cbConfidential.TabIndex = 33;
            this.cbConfidential.Text = "Print \"Confidential\" on each page";
            // 
            // label18
            // 
            this.label18.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label18.Location = new System.Drawing.Point(168, 48);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(88, 16);
            this.label18.TabIndex = 32;
            this.label18.Text = "Sheet to print:";
            // 
            // cbSheet
            // 
            this.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSheet.Location = new System.Drawing.Point(256, 43);
            this.cbSheet.Name = "cbSheet";
            this.cbSheet.Size = new System.Drawing.Size(160, 21);
            this.cbSheet.TabIndex = 31;
            this.cbSheet.SelectedIndexChanged += new System.EventHandler(this.cbSheet_SelectedIndexChanged);
            // 
            // panel3
            // 
            this.panel3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.edBottom);
            this.panel3.Controls.Add(this.label17);
            this.panel3.Controls.Add(this.edRight);
            this.panel3.Controls.Add(this.label16);
            this.panel3.Controls.Add(this.edLeft);
            this.panel3.Controls.Add(this.label15);
            this.panel3.Controls.Add(this.edTop);
            this.panel3.Controls.Add(this.label14);
            this.panel3.Controls.Add(this.label13);
            this.panel3.Controls.Add(this.label12);
            this.panel3.Location = new System.Drawing.Point(504, 232);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(216, 224);
            this.panel3.TabIndex = 30;
            // 
            // edBottom
            // 
            this.edBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edBottom.Location = new System.Drawing.Point(80, 136);
            this.edBottom.Name = "edBottom";
            this.edBottom.Size = new System.Drawing.Size(48, 20);
            this.edBottom.TabIndex = 26;
            this.edBottom.Text = "0";
            // 
            // label17
            // 
            this.label17.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label17.Location = new System.Drawing.Point(16, 160);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(56, 16);
            this.label17.TabIndex = 25;
            this.label17.Text = "Last Col:";
            // 
            // edRight
            // 
            this.edRight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edRight.Location = new System.Drawing.Point(80, 160);
            this.edRight.Name = "edRight";
            this.edRight.Size = new System.Drawing.Size(48, 20);
            this.edRight.TabIndex = 24;
            this.edRight.Text = "0";
            // 
            // label16
            // 
            this.label16.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.Location = new System.Drawing.Point(16, 136);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(85, 16);
            this.label16.TabIndex = 23;
            this.label16.Text = "Last Row:";
            // 
            // edLeft
            // 
            this.edLeft.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edLeft.Location = new System.Drawing.Point(80, 112);
            this.edLeft.Name = "edLeft";
            this.edLeft.Size = new System.Drawing.Size(48, 20);
            this.edLeft.TabIndex = 22;
            this.edLeft.Text = "0";
            // 
            // label15
            // 
            this.label15.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.Location = new System.Drawing.Point(16, 112);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(85, 16);
            this.label15.TabIndex = 21;
            this.label15.Text = "First Col:";
            // 
            // edTop
            // 
            this.edTop.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edTop.Location = new System.Drawing.Point(80, 88);
            this.edTop.Name = "edTop";
            this.edTop.Size = new System.Drawing.Size(48, 20);
            this.edTop.TabIndex = 20;
            this.edTop.Text = "0";
            // 
            // label14
            // 
            this.label14.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(8, 88);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(85, 16);
            this.label14.TabIndex = 3;
            this.label14.Text = "First Row:";
            // 
            // label13
            // 
            this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.label13.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(8, 32);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(192, 32);
            this.label13.TabIndex = 2;
            this.label13.Text = "If one of this values is <=0 all print_range will be printed";
            // 
            // label12
            // 
            this.label12.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(8, 16);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(192, 16);
            this.label12.TabIndex = 1;
            this.label12.Text = "Range to Print:";
            // 
            // Landscape
            // 
            this.Landscape.Location = new System.Drawing.Point(456, 136);
            this.Landscape.Name = "Landscape";
            this.Landscape.Size = new System.Drawing.Size(96, 24);
            this.Landscape.TabIndex = 29;
            this.Landscape.Text = "Landscape";
            // 
            // label11
            // 
            this.label11.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(264, 416);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(80, 16);
            this.label11.TabIndex = 28;
            this.label11.Text = "Footer Margin";
            // 
            // edf
            // 
            this.edf.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edf.Location = new System.Drawing.Point(344, 416);
            this.edf.Name = "edf";
            this.edf.Size = new System.Drawing.Size(128, 20);
            this.edf.TabIndex = 27;
            // 
            // labelb
            // 
            this.labelb.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelb.Location = new System.Drawing.Point(256, 368);
            this.labelb.Name = "labelb";
            this.labelb.Size = new System.Drawing.Size(88, 16);
            this.labelb.TabIndex = 26;
            this.labelb.Text = "Bottom Margin";
            // 
            // edb
            // 
            this.edb.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edb.Location = new System.Drawing.Point(344, 368);
            this.edb.Name = "edb";
            this.edb.Size = new System.Drawing.Size(128, 20);
            this.edb.TabIndex = 25;
            // 
            // label9
            // 
            this.label9.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(56, 368);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(80, 16);
            this.label9.TabIndex = 24;
            this.label9.Text = "Right Margin";
            // 
            // edr
            // 
            this.edr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edr.Location = new System.Drawing.Point(136, 368);
            this.edr.Name = "edr";
            this.edr.Size = new System.Drawing.Size(112, 20);
            this.edr.TabIndex = 23;
            // 
            // label8
            // 
            this.label8.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(264, 328);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(80, 16);
            this.label8.TabIndex = 22;
            this.label8.Text = "Top Margin";
            // 
            // edt
            // 
            this.edt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edt.Location = new System.Drawing.Point(344, 328);
            this.edt.Name = "edt";
            this.edt.Size = new System.Drawing.Size(128, 20);
            this.edt.TabIndex = 21;
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(56, 328);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(80, 16);
            this.label7.TabIndex = 20;
            this.label7.Text = "Left Margin";
            // 
            // edl
            // 
            this.edl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edl.Location = new System.Drawing.Point(136, 328);
            this.edl.Name = "edl";
            this.edl.Size = new System.Drawing.Size(112, 20);
            this.edl.TabIndex = 19;
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(120, 280);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 16);
            this.label4.TabIndex = 18;
            this.label4.Text = "Zoom (%)";
            // 
            // edZoom
            // 
            this.edZoom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edZoom.Location = new System.Drawing.Point(184, 280);
            this.edZoom.Name = "edZoom";
            this.edZoom.Size = new System.Drawing.Size(24, 20);
            this.edZoom.TabIndex = 17;
            // 
            // chFitIn
            // 
            this.chFitIn.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chFitIn.Location = new System.Drawing.Point(56, 248);
            this.chFitIn.Name = "chFitIn";
            this.chFitIn.Size = new System.Drawing.Size(56, 24);
            this.chFitIn.TabIndex = 16;
            this.chFitIn.Text = "Fit in";
            this.chFitIn.CheckedChanged += new System.EventHandler(this.chFitIn_CheckedChanged);
            // 
            // chPrintLeft
            // 
            this.chPrintLeft.Location = new System.Drawing.Point(312, 136);
            this.chPrintLeft.Name = "chPrintLeft";
            this.chPrintLeft.Size = new System.Drawing.Size(136, 24);
            this.chPrintLeft.TabIndex = 15;
            this.chPrintLeft.Text = "Print Left, then down.";
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(256, 248);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 16);
            this.label6.TabIndex = 14;
            this.label6.Text = "pages tall.";
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(144, 248);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(80, 16);
            this.label5.TabIndex = 13;
            this.label5.Text = "pages wide x";
            // 
            // edVPages
            // 
            this.edVPages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edVPages.Location = new System.Drawing.Point(224, 248);
            this.edVPages.Name = "edVPages";
            this.edVPages.ReadOnly = true;
            this.edVPages.Size = new System.Drawing.Size(24, 20);
            this.edVPages.TabIndex = 12;
            // 
            // edHPages
            // 
            this.edHPages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edHPages.Location = new System.Drawing.Point(112, 248);
            this.edHPages.Name = "edHPages";
            this.edHPages.ReadOnly = true;
            this.edHPages.Size = new System.Drawing.Size(24, 20);
            this.edHPages.TabIndex = 10;
            // 
            // edFooter
            // 
            this.edFooter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.edFooter.BackColor = System.Drawing.Color.White;
            this.edFooter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edFooter.Location = new System.Drawing.Point(112, 200);
            this.edFooter.Name = "edFooter";
            this.edFooter.Size = new System.Drawing.Size(608, 20);
            this.edFooter.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(56, 200);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 16);
            this.label3.TabIndex = 7;
            this.label3.Text = "Footer:";
            // 
            // edHeader
            // 
            this.edHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.edHeader.BackColor = System.Drawing.Color.White;
            this.edHeader.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edHeader.Location = new System.Drawing.Point(112, 176);
            this.edHeader.Name = "edHeader";
            this.edHeader.Size = new System.Drawing.Size(608, 20);
            this.edHeader.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(56, 176);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 16);
            this.label2.TabIndex = 5;
            this.label2.Text = "Header:";
            // 
            // edFileName
            // 
            this.edFileName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.edFileName.BackColor = System.Drawing.Color.White;
            this.edFileName.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.edFileName.Location = new System.Drawing.Point(112, 16);
            this.edFileName.Name = "edFileName";
            this.edFileName.ReadOnly = true;
            this.edFileName.Size = new System.Drawing.Size(632, 13);
            this.edFileName.TabIndex = 4;
            this.edFileName.Text = "No file selected";
            this.edFileName.Visible = false;
            // 
            // chFormulaText
            // 
            this.chFormulaText.Location = new System.Drawing.Point(576, 136);
            this.chFormulaText.Name = "chFormulaText";
            this.chFormulaText.Size = new System.Drawing.Size(136, 24);
            this.chFormulaText.TabIndex = 3;
            this.chFormulaText.Text = "Print Formula Text";
            // 
            // chGridLines
            // 
            this.chGridLines.Location = new System.Drawing.Point(56, 136);
            this.chGridLines.Name = "chGridLines";
            this.chGridLines.Size = new System.Drawing.Size(104, 24);
            this.chGridLines.TabIndex = 2;
            this.chGridLines.Text = "Print Grid Lines";
            // 
            // chAntiAlias
            // 
            this.chAntiAlias.Location = new System.Drawing.Point(56, 88);
            this.chAntiAlias.Name = "chAntiAlias";
            this.chAntiAlias.Size = new System.Drawing.Size(152, 16);
            this.chAntiAlias.TabIndex = 1;
            this.chAntiAlias.Text = "Antialias Text";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(24, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "File to print:";
            this.label1.Visible = false;
            // 
            // label10
            // 
            this.label10.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(48, 416);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(88, 16);
            this.label10.TabIndex = 22;
            this.label10.Text = "Header Margin";
            // 
            // edh
            // 
            this.edh.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edh.Location = new System.Drawing.Point(136, 416);
            this.edh.Name = "edh";
            this.edh.Size = new System.Drawing.Size(112, 20);
            this.edh.TabIndex = 21;
            // 
            // exportImageDialog
            // 
            this.exportImageDialog.DefaultExt = "png";
            this.exportImageDialog.Filter = "Png files|*.png|Jpg files|*.jpg";
            this.exportImageDialog.Title = "Save image as...";
            // 
            // exportTiffDialog
            // 
            this.exportTiffDialog.DefaultExt = "tif";
            this.exportTiffDialog.Filter = "TIFF Files|*.tif";
            this.exportTiffDialog.Title = "Save image as multi page tiff...";
            // 
            // exportImagesMenu
            // 
            this.exportImagesMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.ExportUsingPrintController,
            this.menuItem6,
            this.ExportUsingFlexCelImgExport,
            this.ExportUsingSaveAsImage,
            this.ExportMultiPageTiff,
            this.menuItem1});
            // 
            // ExportUsingPrintController
            // 
            this.ExportUsingPrintController.Index = 0;
            this.ExportUsingPrintController.Text = "Using Printcontroller (old way)";
            this.ExportUsingPrintController.Click += new System.EventHandler(this.ExportUsingPrintController_Click);
            // 
            // menuItem6
            // 
            this.menuItem6.Index = 1;
            this.menuItem6.Text = "-";
            // 
            // ExportUsingFlexCelImgExport
            // 
            this.ExportUsingFlexCelImgExport.Index = 2;
            this.ExportUsingFlexCelImgExport.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.ImgBlackAndWhite,
            this.Img256Colors,
            this.ImgTrueColor});
            this.ExportUsingFlexCelImgExport.Text = "Using FlexCelImgExport -All pages (recommended way)";
            // 
            // ImgBlackAndWhite
            // 
            this.ImgBlackAndWhite.Index = 0;
            this.ImgBlackAndWhite.Text = "Black And White";
            this.ImgBlackAndWhite.Click += new System.EventHandler(this.ImgBlackAndWhite_Click);
            // 
            // Img256Colors
            // 
            this.Img256Colors.Index = 1;
            this.Img256Colors.Text = "256 Colors";
            this.Img256Colors.Click += new System.EventHandler(this.Img256Colors_Click);
            // 
            // ImgTrueColor
            // 
            this.ImgTrueColor.Index = 2;
            this.ImgTrueColor.Text = "True Color";
            this.ImgTrueColor.Click += new System.EventHandler(this.ImgTrueColor_Click);
            // 
            // ExportUsingSaveAsImage
            // 
            this.ExportUsingSaveAsImage.Index = 3;
            this.ExportUsingSaveAsImage.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.ImgBlackAndWhite2,
            this.Img256Colors2,
            this.ImgTrueColor2});
            this.ExportUsingSaveAsImage.Text = "Using FlexCelImgExport  - 1 page (recommended way)";
            // 
            // ImgBlackAndWhite2
            // 
            this.ImgBlackAndWhite2.Index = 0;
            this.ImgBlackAndWhite2.Text = "Black And White";
            this.ImgBlackAndWhite2.Click += new System.EventHandler(this.ImgBlackAndWhite2_Click);
            // 
            // Img256Colors2
            // 
            this.Img256Colors2.Index = 1;
            this.Img256Colors2.Text = "256 Colors";
            this.Img256Colors2.Click += new System.EventHandler(this.Img256Colors2_Click);
            // 
            // ImgTrueColor2
            // 
            this.ImgTrueColor2.Index = 2;
            this.ImgTrueColor2.Text = "True Color";
            this.ImgTrueColor2.Click += new System.EventHandler(this.ImgTrueColor2_Click);
            // 
            // ExportMultiPageTiff
            // 
            this.ExportMultiPageTiff.Index = 4;
            this.ExportMultiPageTiff.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.TiffFax,
            this.TiffBlackAndWhite,
            this.Tiff256Colors,
            this.TiffTrueColor});
            this.ExportMultiPageTiff.Text = "MultiPage TIFF using FlexCelImgExport";
            // 
            // TiffFax
            // 
            this.TiffFax.Index = 0;
            this.TiffFax.Text = "Fax";
            this.TiffFax.Click += new System.EventHandler(this.TiffFax_Click);
            // 
            // TiffBlackAndWhite
            // 
            this.TiffBlackAndWhite.Index = 1;
            this.TiffBlackAndWhite.Text = "Black And White";
            this.TiffBlackAndWhite.Click += new System.EventHandler(this.TiffBlackAndWhite_Click);
            // 
            // Tiff256Colors
            // 
            this.Tiff256Colors.Index = 2;
            this.Tiff256Colors.Text = "256 Colors";
            this.Tiff256Colors.Click += new System.EventHandler(this.Tiff256Colors_Click);
            // 
            // TiffTrueColor
            // 
            this.TiffTrueColor.Index = 3;
            this.TiffTrueColor.Text = "True Color";
            this.TiffTrueColor.Click += new System.EventHandler(this.TiffTrueColor_Click);
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 5;
            this.menuItem1.Text = "-";
            // 
            // printDialog1
            // 
            this.printDialog1.Document = this.flexCelPrintDocument1;
            this.printDialog1.UseEXDialog = true;
            // 
            // printPreviewDialog1
            // 
            this.printPreviewDialog1.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.ClientSize = new System.Drawing.Size(400, 300);
            this.printPreviewDialog1.Document = this.flexCelPrintDocument1;
            this.printPreviewDialog1.Enabled = true;
            this.printPreviewDialog1.Icon = ((System.Drawing.Icon)(resources.GetObject("printPreviewDialog1.Icon")));
            this.printPreviewDialog1.Name = "printPreviewDialog1";
            this.printPreviewDialog1.Visible = false;
            // 
            // frmFlexPrintPreview
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(768, 517);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Name = "frmFlexPrintPreview";
            this.Text = "Print and preview";
            this.Load += new System.EventHandler(this.frmPrintPreview_Load);
            this.panel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

        private Button preview;
        private PrintDialog printDialog1;
        private PrintPreviewDialog printPreviewDialog1;

    }
}

