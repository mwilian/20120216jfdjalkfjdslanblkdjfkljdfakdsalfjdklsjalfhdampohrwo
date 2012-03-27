using System;
using System.ComponentModel;
using System.IO;
using FlexCel.Pdf;
using FlexCel.Core;

#if (WPF)
using RectangleF = System.Windows.Rect;
#else
using System.Drawing;
#endif

namespace FlexCel.Render
{
    /// <summary>
    /// A component for exporting an Excel file to PDF.
    /// </summary>
    public class FlexCelPdfExport : Component
    {
        #region Privates
        private FlexCelRender FRenderer = null;

        /// <summary>
        /// Stream where the file will be saved. You can use this in a derived class to write your own data to the stream.
        /// </summary>
		protected Stream FPdfStream = null;

        private volatile bool FCanceled;
        internal volatile FlexCelPdfExportProgress FProgress;

        /// <summary>
        /// If true, the file will be compressed. This property is just for derived classes, for normal cases use <see cref="Compress"/>
        /// </summary>
        protected bool FCompress;

        bool FKerning;
        bool FAllowOverwritingFiles;
        TPageLayout FPageLayout;

        /// <summary>
        /// Only for use in derived classes. Use <see cref="Properties"/> instead.
        /// </summary>
        protected TPdfProperties FProperties;

        /// <summary>
        /// Writer where the pdf commands will be sent. Only for use in derived classes.
        /// </summary>
        protected PdfWriter PdfCanvas;

        bool FirstPage;

        int CurrentTotalPage;
        int FirstPageInSheet;

        /// <summary>
        /// Only for use in derived classes. Use <see cref="FPrintRange"/> instead.
        /// </summary>
        protected TXlsCellRange FPrintRange;

        /// <summary>
        /// Only for use in derived classes. Use <see cref="PageSize"/> instead.
        /// </summary>
        protected TPaperDimensions FPageSize;

        /// <summary>
        /// Only for use in derived classes. Use <see cref="FontEmbed"/> instead.
        /// </summary>
        protected TFontEmbed FFontEmbed;

        /// <summary>
        /// Only for use in derived classes. Use <see cref="FontSubset"/> instead.
        /// </summary>
        protected TFontSubset FFontSubset;

        /// <summary>
        /// Only for use in derived classes. Use <see cref="FontMapping"/> instead.
        /// </summary>
        protected TFontMapping FFontMapping;

        /// <summary>
        /// Only for use in derived classes. Use <see cref="FallbackFonts"/> instead.
        /// </summary>
        protected string FFallbackFonts;

        TPdfSignature FSignature;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new FlexCelPdfExport instance.
        /// </summary>
        public FlexCelPdfExport()
        {
            FCompress = true;
            FRenderer = new FlexCelRender();
            FRenderer.ReverseRightToLeftStrings = true;
            FPrintRange = new TXlsCellRange(0, 0, 0, 0);
            FProperties = new TPdfProperties();
            PageSize = null;
            FProgress = new FlexCelPdfExportProgress();

            FFontEmbed = TFontEmbed.None;
			FFontSubset = TFontSubset.Subset;
            FFontMapping = TFontMapping.ReplaceStandardFonts;
			FFallbackFonts = "Arial Unicode MS";

        }

        /// <summary>
        /// Creates a new FlexCelPdfExport and assigns it to an ExcelFile.
        /// </summary>
        /// <param name="aWorkbook">ExcelFile containing the data this component will export.</param>
        public FlexCelPdfExport(ExcelFile aWorkbook)
            : this()
        {
            FRenderer.Workbook = aWorkbook;
        }

        /// <summary>
        /// Creates a new FlexCelPdfExport and assigns it to an ExcelFile, setting AllowOverwritingFiles to the desired value.
        /// </summary>
        /// <param name="aWorkbook">ExcelFile containing the data this component will export.</param>
        /// <param name="aAllowOverwritingFiles">When true, existing files will be overwrited.</param>
        public FlexCelPdfExport(ExcelFile aWorkbook, bool aAllowOverwritingFiles)
            : this()
        {
            FRenderer.Workbook = aWorkbook;
            FAllowOverwritingFiles = aAllowOverwritingFiles;
        }

        #endregion

        #region Properties
        /// <summary>
        /// If true the export has been canceled with <see cref="Cancel"/> method.
        /// You can't set this variable to false, and setting it true is the same as calling <see cref="Cancel"/>.
        /// </summary>
        [Browsable(false),
        DefaultValue(false)]
        public bool Canceled
        {
            get { return FCanceled; }
            set
            {
                if (value == true) FCanceled = true; //Don't allow to uncancel.
            }
        }

        /// <summary>
        /// Progress of the export. This variable must be accessed from other thread.
        /// </summary>
        [Browsable(false)]
        public FlexCelPdfExportProgress Progress
        {
            get { return FProgress; }
        }

        /// <summary>
        /// The ExcelFile to print.
        /// </summary>
        [Browsable(false)]
        public virtual ExcelFile Workbook { get { return FRenderer.Workbook; } set { FRenderer.Workbook = value; } }

        /// <summary>
        /// Select which kind of objects should not be printed or exported to pdf.
        /// </summary>
        [Category("Behavior"),
        Description("Select which kind of objects should not be printed or exported to pdf."),
        DefaultValue(THidePrintObjects.None)]
        public virtual THidePrintObjects HidePrintObjects { get { return FRenderer.HidePrintObjects; } set { FRenderer.HidePrintObjects = value; } }

        /// <summary>
        /// When true, the pdf file will be compressed.
        /// </summary>
        [Category("Behavior"),
        Description("When true, the pdf file will be compressed."),
        DefaultValue(true)]
        public bool Compress { get { return FCompress; } set { FCompress = value; } }

        /// <summary>
        /// By default, pdf does not do any kerning with the fonts. This is, on the string "AVANT", it won't
        /// compensate the spaces between "A" and "V". (they should be smaller) 
        /// If you turn this property on, FlexCel will calculate the kerning and add it to the generated file.
        /// The result file will be a little bigger because of the kerning info on all strings, but it will also
        /// look a little better.
        /// </summary>
        [Category("Behavior"),
        Description("If you turn this property on, FlexCel will calculate the kerning and add it to the generated file."),
        DefaultValue(false)]
        public bool Kerning { get { return FKerning; } set { FKerning = value; } }

        /// <summary>
        /// Pdf file properties.
        /// </summary>
        [Category("Properties"),
        Description("Pdf file properties.")]
        public TPdfProperties Properties { get { return FProperties; } set { if (value == null) FProperties = new TPdfProperties(); else FProperties = value; } }

        /// <summary>
        /// First column to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated.
        /// </summary>
        [Category("PrintRange"),
        Description("First column to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated."),
        DefaultValue(0)]
        public int PrintRangeLeft { get { return FPrintRange.Left; } set { FPrintRange.Left = value; } }

        /// <summary>
        /// First row to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated. 
        /// </summary>
        [Category("PrintRange"),
        Description("First row to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated."),
        DefaultValue(0)]
        public int PrintRangeTop { get { return FPrintRange.Top; } set { FPrintRange.Top = value; } }

        /// <summary>
        /// Last column to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated.
        /// </summary>
        [Category("PrintRange"),
        Description("Last column to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated."),
        DefaultValue(0)]
        public int PrintRangeRight { get { return FPrintRange.Right; } set { FPrintRange.Right = value; } }

        /// <summary>
        /// Last row to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated.
        /// </summary>
        [Category("PrintRange"),
        Description("Last row to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated."),
        DefaultValue(0)]
        public int PrintRangeBottom { get { return FPrintRange.Bottom; } set { FPrintRange.Bottom = value; } }

        /// <summary>
        /// Pdf page size. Set it to null to use the paper size on the xls file.
        /// </summary>
        public TPaperDimensions PageSize
        {
            get { return FPageSize; }
            set
            {
                FPageSize = value;
            }
        }

        /// <summary>
        /// Determines if FlexCel will automatically delete existing pdf files or not.
        /// </summary>
        [Category("Behavior"),
        Description("Determines if FlexCel will automatically delete existing pdf files or not."),
        DefaultValue(false)]
        public bool AllowOverwritingFiles { get { return FAllowOverwritingFiles; } set { FAllowOverwritingFiles = value; } }

		/// <summary>
		/// Determines what fonts will be embedded on the generated pdf. Note that when using UNICODE, fonts will be embedded anyway, no matter what this setting is.
		/// </summary>
		[Category("Fonts"),
		Description("Determines what fonts will be embedded on the generated pdf. Note that " +
			"when using UNICODE, fonts will be embedded anyway, no matter what this setting is."),
		DefaultValue(TFontEmbed.None)]
		public TFontEmbed FontEmbed { get { return FFontEmbed; } set { FFontEmbed = value; } }

		/// <summary>
		/// Determines if the full font will be embedded or only the characters used, when embedding fonts.
		/// </summary>
		[Category("Fonts"),
		Description("Determines if the full font will be embedded or only the characters used, when embedding fonts."),
		DefaultValue(TFontSubset.Subset)]
		public TFontSubset FontSubset { get { return FFontSubset; } set { FFontSubset = value; } }

        /// <summary>
        /// Determines how fonts will be replaced on the generated pdf. Pdf comes with 4 standard font families,
        /// Serif, Sans-Serif, Monospace and Symbol. You can use for example the standard Helvetica instead of Arial and do not worry about embedding the font.
        /// </summary>
        [Category("Fonts"),
        Description("Determines how fonts will be replaced on the generated pdf. Pdf comes with 4 standard font families," +
          "Serif, Sans-Serif, Monospace and Symbol. You can use for example the standard Helvetica instead of Arial and do not worry about embedding the font."),
        DefaultValue(TFontMapping.ReplaceStandardFonts)]
        public TFontMapping FontMapping { get { return FFontMapping; } set { FFontMapping = value; } }

		/// <summary>
		/// A semicolon (;) separated list of font names to try when a character is not found in the used font.<br/>
		/// When a character is not found in a font, it will display as an empty square by default. By setting this
		/// property, FlexCel will try to find a font that supports this character in this list, and if found, use that font
		/// to render the character.
		/// </summary>
		[Category("Fonts"),
		Description("A semicolon (;) separated list of font names to try when a character is not found in the used font."),
		DefaultValue("Arial Unicode MS")]
		public string FallbackFonts { get { return FFallbackFonts; } set { FFallbackFonts = value; } }


        /// <summary>
        /// Sets the default page layout when opening the document.
        /// </summary>
        public TPageLayout PageLayout { get { return FPageLayout; } set { FPageLayout = value; } }


        /// <summary>
        /// Returns the next page that we are going to print.
        /// </summary>
        public int CurrentPage { get { return FirstPage ? 1 : CurrentTotalPage + 1; } }

        /// <summary>
        /// Returns the next page we are going to print, on the current sheet.
        /// When not printing more than one sheet, it is equivalent to <see cref="CurrentPage"/>
        /// </summary>
        public int CurrentPageInSheet { get { return CurrentPage - FirstPageInSheet; } }

        #endregion

        #region Events
        /// <summary>
        /// Fires before each new page is generated on the pdf.
        /// You can use this event to change the pagesize for the new sheet.
        /// </summary>
        [Category("Generate"),
        Description("Fires before each new page is generated on the pdf. You can use this event to change the pagesize for the new sheet.")]
        public event PageEventHandler BeforeNewPage;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelPdfExport. See also <see cref="BeforeNewPage"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnBeforeNewPage(PageEventArgs e)
        {
            if (BeforeNewPage != null) BeforeNewPage(this, e);
        }

        /// <summary>
        /// Fires after each new page is generated on the pdf, but before any content is written to the page. (The page is blank)
        /// You can use this event to add a watermark or a background image.
        /// </summary>
        [Category("Generate"),
        Description("Fires after each new page is generated on the pdf, but before any content is written to the page. (The page is blank).You can use this event to add a watermark or a background image.")]
        public event PageEventHandler BeforeGeneratePage;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelPdfExport. See also <see cref="BeforeGeneratePage"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnBeforeGeneratePage(PageEventArgs e)
        {
            if (BeforeGeneratePage != null) BeforeGeneratePage(this, e);
        }

        /// <summary>
        /// Fires after each new page is generated on the pdf, and after all content is written to the page. (The page is written)
        /// You can use this event to add some text or images on top of the page contents.
        /// </summary>
        [Category("Generate"),
        Description("Fires after each new page is generated on the pdf, and after all content is written to the page. (The page is written).You can use this event to add some text or images on top of the page contents.")]
        public event PageEventHandler AfterGeneratePage;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelPdfExport. See also <see cref="AfterGeneratePage"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnAfterGeneratePage(PageEventArgs e)
        {
            if (AfterGeneratePage != null) AfterGeneratePage(this, e);
        }

        /// <summary>
        /// Use this event if you want to provide your own font information for embedding. 
        /// Note that if you don't assign this event, the default method will be used, and this 
        /// will try to find the font on the Fonts folder. To change the font folder, use <see cref="GetFontFolder"/> event
        /// </summary>
        [Category("Fonts"),
		Description("Use this event if you want to provide your own font information for embedding.")]
        public event GetFontDataEventHandler GetFontData;

        /// <summary>
        /// Use this event if you want to provide your own font information for embedding. 
        /// Normally FlexCel will search for fonts on [System]\Fonts folder. If your fonts are in 
        /// other location, you can tell FlexCel where they are here. If you prefer just to give FlexCel
        /// the full data on the font, you can use <see cref="GetFontData"/> event instead.
        /// </summary>
		[Category("Fonts"),
		Description("Use this event if you want to provide your own font information for embedding.")]
		public event GetFontFolderEventHandler GetFontFolder;

		/// <summary>
		/// Use this event if you want to manually specify which fonts to embed into the pdf document.
		/// </summary>
		[Category("Fonts"),
		Description("Use this event if you want to manually specify which fonts to embed into the pdf document.")]
		public event FontEmbedEventHandler OnFontEmbed;

        /// <summary>
        /// Use this event to customize what goes inside the bookmarks when exporting multiple sheets of an xls file.
        /// </summary>
        [Category("Generate"),
        Description("Use this event to customize what goes inside the bookmarks when exporting multiple sheets of an xls file.")]
        public event GetBookmarkInformationEventHandler GetBookmarkInformation;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelPdfExport.
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnGetBookmarkInformation(GetBookmarkInformationArgs e)
        {
            if (GetBookmarkInformation != null) GetBookmarkInformation(this, e);
        }


        #endregion

        #region Thread Methods
        /// <summary>
        /// Cancels a running export. This method is equivalent to setting <see cref="Canceled"/> = true.
        /// </summary>
        public void Cancel()
        {
            Canceled = true;
        }

        #endregion

        #region Export
        /// <summary>
        /// Exports the active sheet of the associated xls workbook to a stream.
        /// </summary>
        /// <remarks>
        /// This method is a shortcut for calling <see cref="BeginExport(Stream)"/>/ <see cref="ExportSheet()"/>/<see cref="EndExport()"/>.
        /// </remarks>
        /// <param name="pdfStream">Stream where the result will be written.</param>
        public void Export(Stream pdfStream)
        {
            BeginExport(pdfStream);
            ExportSheet();
            EndExport();
        }

        /// <summary>
        /// Exports the active sheet of the the associated xls workbook to a file.
        /// </summary>
        /// <remarks>
        /// This method is a shortcut for calling <see cref="BeginExport(Stream)"/>/ <see cref="ExportSheet()"/>/<see cref="EndExport()"/>.
        /// </remarks>
        /// <param name="fileName">File to export.</param>
        public void Export(string fileName)
        {
            try
            {
                FileMode fm = FileMode.CreateNew;
                if (AllowOverwritingFiles) fm = FileMode.Create;
                using (FileStream f = new FileStream(fileName, fm, FileAccess.Write))
                {
                    Export(f);
                }
            }
            catch (IOException)
            {
                //Don't delete the file in an io exception. It might be because allowoverwritefiles was false, and the file existed.
                throw;
            }
            catch
            {
                FlxUtils.TryDelete(fileName);
                throw;
            }
        }

        private void LoadPageSize()
        {
            TPaperDimensions pd = FPageSize;
            if (FPageSize == null)
            {
                pd = Workbook.PrintPaperDimensions;
                if ((Workbook.PrintOptions & TPrintOptions.Orientation) == 0)
                {
                    float w = pd.Width;
                    pd.Width = pd.Height;
                    pd.Height = w;
                }
            }
            PdfCanvas.PageSize = pd;
        }

        /// <summary>
        /// Initializes the PDF exporting to a new file. After calling this method
        /// you can call <see cref="ExportSheet()"/> to export different xls files to the same pdf.
        /// You should always end the document with a call to <see cref="EndExport"/>
        /// </summary>
        /// <param name="pdfStream">Stream that will contain the new pdf file.</param>
        public void BeginExport(Stream pdfStream)
        {
            FCanceled = false;
            PdfCanvas = new PdfWriter();
            PdfCanvas.Sign(FSignature);
            PdfCanvas.Compress = FCompress;
            PdfCanvas.Kerning = FKerning;
            PdfCanvas.Properties = FProperties;
            PdfCanvas.GetFontData = GetFontData;
            PdfCanvas.GetFontFolder = GetFontFolder;
			PdfCanvas.OnFontEmbed = OnFontEmbed;
			PdfCanvas.FallbackFonts = FFallbackFonts;
            CurrentTotalPage = 1;
            FirstPageInSheet = 0;
            FPdfStream = pdfStream;

            FRenderer.CreateFontCache();

            FirstPage = true;
            FProgress.Clear();
        }

        /// <summary>
        /// Writes the trailer information on a PDF file. Always call this method after calling <see cref="BeginExport"/>
        /// </summary>
        public void EndExport()
        {
            PdfCanvas.EndDoc();
            FRenderer.DisposeFontCache();
            FPdfStream = null;
            PdfCanvas = null;
        }

        private void PrepareCanvas()
        {
            PdfCanvas.YAxisGrowsDown = true;
            PdfCanvas.Scale = 72F / FlexCelRender.DispMul;
            PdfCanvas.AddFontDescent = true;
            PdfCanvas.FontEmbed = FontEmbed;
			PdfCanvas.FontSubset = FontSubset;
            PdfCanvas.FontMapping = FontMapping;
			PdfCanvas.FallbackFonts = FFallbackFonts;
        }

        /// <summary>
        /// Calculates the actual spreadsheet range that will be printed. This is given by:
        /// 1)If you specified non zero values on PrintRange, this will be used.
        /// 2)If any value in PrintRange is zero and there is a Print Area defined on the
        /// spreadsheet, the Print Area will be used.
        /// 3)If there is no PrintRange and no Print Area defined, the visible cells on the
        /// sheet will be printed.
        /// </summary>
        ///<returns>The area that will be exported.</returns>
        public TXlsCellRange CalcPrintArea()
        {
			Workbook.Recalc(false);
            PdfCanvas = new PdfWriter();
            PdfGraphics RealCanvas = new PdfGraphics(PdfCanvas);
            PrepareCanvas();
            IFlxGraphics aCanvas = RealCanvas;
            FRenderer.SetCanvas(aCanvas);
            try
            {
                FRenderer.CreateFontCache();
                try
                {
                    return FRenderer.InternalCalcPrintArea(FPrintRange)[0];
                }
                finally
                {
                    FRenderer.DisposeFontCache();
                }
            }
            finally
            {
                FRenderer.SetCanvas(null);
            }
        }

        /// <summary>
        /// Returns the number of pages that the active sheet will use when exported to pdf.
        /// </summary>
        public int TotalPagesInSheet()
        {
            int PagesInSheet;
            IFlxGraphics aCanvas = new PdfGraphics(PdfCanvas);
            LoadPageSize();

            PrepareCanvas();
            FRenderer.SetCanvas(aCanvas);
            try
            {
                TXlsCellRange[] MyPrintRange = FRenderer.InternalCalcPrintArea(FPrintRange);
                TXlsCellRange PagePrintRange;
                RectangleF[] PaintClipRect;

                FRenderer.InitializePrint(aCanvas, PdfGraphics.ConvertToUnits(PdfCanvas.PageSize),
                    PdfGraphics.ConvertToUnits(PdfCanvas.PageSize),
                    MyPrintRange, out PaintClipRect, out PagesInSheet, out PagePrintRange);


                return PagesInSheet;
            }
            finally
            {
                FRenderer.SetCanvas(null);
            }
        }

        /// <summary>
        /// Exports the activesheet on the current XlsFile.
        /// </summary>
        public void ExportSheet()
        {
            ExportSheet(1, -1);
        }

        /// <summary>
        /// Exports the active sheet on the current XlsFile. You can define which is the first page to print and the global count of pages, so the 
        /// page numbers on headers and footers of the excel file correspond with the actual pages on the pdf.
        /// </summary>
        /// <param name="startPage">Fist page that the headers and footers on the xls file will show. If you are exporting only one sheet to the pdf file,
        /// this can be 1. If you are exporting more than one sheet to the same pdf file, you will want to set StartPage to the actual page on the pdf.</param>
        /// <param name="totalPages">The total number of pages to display on Excel headers and footers. If you are exporting only one sheet to the pdf file, set it to -1, and it will be calculated automatically. If not, please suply here the total number of pages the file will have so FlexCel can show footers like "page 1 of 50"</param>
        public void ExportSheet(int startPage, int totalPages)
        {
            if (PdfCanvas == null) FlxMessages.ThrowException(FlxErr.ErrBeginExportNotCalled);
            Workbook.Recalc(false);

            FirstPageInSheet = startPage - 1;

            PdfGraphics RealCanvas = new PdfGraphics(PdfCanvas);

            LoadPageSize();
            if (FirstPage)
            {
                OnBeforeNewPage(new PageEventArgs(PdfCanvas, CurrentTotalPage, CurrentTotalPage - FirstPageInSheet));
                PrepareCanvas();
                PdfCanvas.BeginDoc(FPdfStream);
                PdfCanvas.PageLayout = FPageLayout;
                FPdfStream = null;
            }

            IFlxGraphics aCanvas = RealCanvas;
            FRenderer.SetCanvas(aCanvas);
            try
            {
                RectangleF[] PaintClipRect;
                TXlsCellRange[] MyPrintRange = FRenderer.InternalCalcPrintArea(FPrintRange);

                TXlsCellRange PagePrintRange;

                int PagesInSheet;
                FRenderer.InitializePrint(aCanvas, PdfGraphics.ConvertToUnits(PdfCanvas.PageSize),
                    PdfGraphics.ConvertToUnits(PdfCanvas.PageSize),
                    MyPrintRange, out PaintClipRect, out PagesInSheet, out PagePrintRange);

                if (totalPages < 0) totalPages = PagesInSheet;
                FProgress.SetTotalPages(totalPages);

                int PrintArea = 0;
                for (int i = 0; i < PagesInSheet; i++)
                {
                    if (Canceled) return;
                    LoadPageSize();

                    if (!FirstPage)
                    {
                        CurrentTotalPage++;
                        OnBeforeNewPage(new PageEventArgs(PdfCanvas, CurrentTotalPage, CurrentTotalPage - FirstPageInSheet));
                        PrepareCanvas();
                        PdfCanvas.NewPage();
                    }
                    OnBeforeGeneratePage(new PageEventArgs(PdfCanvas, CurrentTotalPage, CurrentTotalPage - FirstPageInSheet));
                    PrepareCanvas();
                    PdfCanvas.SaveState();

                    FirstPage = false;

                    if (Canceled) return;
                    FProgress.SetPage(startPage + i);
                    FRenderer.GenericPrint(aCanvas, PdfGraphics.ConvertToUnits(PdfCanvas.PageSize), MyPrintRange, startPage + i, 
                        PaintClipRect, totalPages, true, PagePrintRange, ref PrintArea);
                    if (Canceled) return;

                    PdfCanvas.RestoreState();
                    OnAfterGeneratePage(new PageEventArgs(PdfCanvas, CurrentTotalPage, CurrentTotalPage - FirstPageInSheet));
                    PrepareCanvas();
                }
            }
            finally
            {
                FRenderer.SetCanvas(null);
            }
        }

        /// <summary>
        /// This method will export all the visible sheets on an xls file to pdf.
        /// Different than calling ExportSheet for each sheet, this method can keep the page number growing on each sheet, without resetting it.
        /// </summary>
        /// <param name="bookmarkName">If not null, each sheet will be added as an entry on the Bookmarks in the pdf file, under the name specified here.
        /// If you want the Bookmark window to be visible when you open the pdf file, set <see cref="PdfWriter.PageLayout"/> = <see cref="TPageLayout"/> 
        /// Also, use the <see cref="GetBookmarkInformation"/> event to further customize what goes in each of the entries.
        /// </param>
        /// <param name="resetPageNumberOnEachSheet">If true, each new sheet will reset the page number shown on Excel headers and footers.</param>
        public virtual void ExportAllVisibleSheets(bool resetPageNumberOnEachSheet, string bookmarkName)
        {
            if (PdfCanvas == null) FlxMessages.ThrowException(FlxErr.ErrBeginExportNotCalled);
            TBookmark ParentBookmark = new TBookmark(bookmarkName, new TPdfDestination(CurrentPage), false);
            OnGetBookmarkInformation(new GetBookmarkInformationArgs(PdfCanvas, 0, 0, ParentBookmark));

			int SaveActiveSheet = Workbook.ActiveSheet;
            try
            {
                int TotalPages = -1;
                if (!resetPageNumberOnEachSheet)
                {
                    TotalPages = 0;
                    //Calculate total pages on all sheets.
                    for (int sheet = 1; sheet <= Workbook.SheetCount; sheet++)
                    {
                        Workbook.ActiveSheet = sheet;
                        if (Workbook.SheetVisible != TXlsSheetVisible.Visible) continue;
                        TotalPages += TotalPagesInSheet();
                    }

                }

                for (int sheet = 1; sheet <= Workbook.SheetCount; sheet++)
                {
                    Workbook.ActiveSheet = sheet;
                    if (Workbook.SheetVisible != TXlsSheetVisible.Visible) continue;

                    int StartSheet = 1;
                    if (!resetPageNumberOnEachSheet)
                    {
                        StartSheet = CurrentPage;
                    }


                    if (bookmarkName != null)
                    {
                        TBookmark bmk = new TBookmark(Workbook.SheetName, new TPdfDestination(CurrentPage), false);
                        OnGetBookmarkInformation(new GetBookmarkInformationArgs(PdfCanvas, CurrentPage, CurrentPageInSheet, bmk));
                        ParentBookmark.AddChild(bmk);
                    }
                    ExportSheet(StartSheet, TotalPages);

                }
                if (bookmarkName != null)
                {
                    PdfCanvas.AddBookmark(ParentBookmark);
                }

            }
            finally
            {
                Workbook.ActiveSheet = SaveActiveSheet;
            }
        }
        #endregion

        #region Sign
        /// <summary>
        /// Signs the pdf documents with the specified <see cref="TPdfSignature"/> or <see cref="TPdfVisibleSignature"/>.
        /// <b>Note:</b> This method must be called <b>before</b> calling <see cref="BeginExport"/>
        /// </summary>
        /// <param name="aSignature">Signature used for signing. Set it to null to stop signing the next documents.</param>
        public void Sign(TPdfSignature aSignature)
        {
            FSignature = aSignature;
        }
        #endregion


    }
}

