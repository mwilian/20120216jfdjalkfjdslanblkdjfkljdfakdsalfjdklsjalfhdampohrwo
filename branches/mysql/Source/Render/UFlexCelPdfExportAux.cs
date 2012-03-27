using System;
using FlexCel.Pdf;

namespace FlexCel.Render
{
    /// <summary>
    /// Indicates how much of the report has been generated.
    /// </summary>
    public class FlexCelPdfExportProgress
    {
        private volatile int FPage;
        private volatile int FTotalPages;

        internal FlexCelPdfExportProgress()
        {
            Clear();
        }

        internal void Clear()
        {
            FPage=0;
            FTotalPages=0;
        }

        internal void SetPage(int value)
        {
            FPage=value;
        }

        internal void SetTotalPages(int value)
        {
            FTotalPages=value;
        }

        /// <summary>
        /// The page that is being written.
        /// </summary>
        public int Page{get {return FPage;}}

        /// <summary>
        /// The total number of pages exporting.
        /// </summary>
        public int TotalPage{get {return FTotalPages;}}
    }

    #region Event Handlers
    /// <summary>
    /// Arguments passed on <see cref="FlexCel.Render.FlexCelPdfExport.BeforeGeneratePage"/>, 
    /// <see cref="FlexCel.Render.FlexCelPdfExport.BeforeNewPage"/> and <see cref="FlexCel.Render.FlexCelPdfExport.AfterGeneratePage"/>
    /// </summary>
    public class PageEventArgs: EventArgs
    {
		private readonly int FCurrentPage;
		private readonly int FCurrentPageInSheet;
        private readonly PdfWriter FPdfWriter;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        /// <param name="aPdfWriter">The file we are processing.</param>
        /// <param name="aCurrentPage">The page that is being generated.</param>
        /// <param name="aCurrentPageInSheet">The page in the current sheet that is being generated.</param>
        public PageEventArgs(PdfWriter aPdfWriter, int aCurrentPage, int aCurrentPageInSheet)
        {
            FPdfWriter=aPdfWriter;
            FCurrentPage=aCurrentPage;
			FCurrentPageInSheet = aCurrentPageInSheet;
        }

        /// <summary>
        /// The file with the pdf data.
        /// </summary>
        public PdfWriter File
        {
            get {return FPdfWriter;}
        }

        /// <summary>
        /// Page currently printing.
        /// </summary>
        public int CurrentPage {get {return FCurrentPage;}}

		/// <summary>
		/// Page currently printing on the sheet printing.
		/// </summary>
		public int CurrentPageInSheet {get {return FCurrentPageInSheet;}}
	}

    /// <summary>
    /// Generic delegate for After/Before page events.
    /// </summary>
    public delegate void PageEventHandler(object sender, PageEventArgs e);

	#region GetBookmarkInformation
	/// <summary>
	/// Arguments passed on <see cref="FlexCel.Render.FlexCelPdfExport.GetBookmarkInformation"/>, 
	/// </summary>
	public class GetBookmarkInformationArgs: EventArgs
	{
		private readonly int FCurrentPage;
		private readonly int FCurrentPageInSheet;
		private readonly PdfWriter FPdfWriter;
		private readonly TBookmark FBookmark;

		/// <summary>
		/// Creates a new Argument.
		/// </summary>
		/// <param name="aPdfWriter">The file we are processing.</param>
		/// <param name="aCurrentPage">The page that is being generated. 0 means the global bookmark parent of all the sheets.</param>
		/// <param name="aCurrentPageInSheet">The page that is being generated, relative to the sheet.</param>
		/// <param name="aBookmark">Bookmark that we are about to include. you can customize it on this event.</param> 
		public GetBookmarkInformationArgs(PdfWriter aPdfWriter, int aCurrentPage, int aCurrentPageInSheet, TBookmark aBookmark)
		{
			FPdfWriter=aPdfWriter;
			FCurrentPage=aCurrentPage;
			FCurrentPageInSheet=aCurrentPageInSheet;
			FBookmark = aBookmark;
		}

		/// <summary>
		/// The file with the pdf data.
		/// </summary>
		public PdfWriter File
		{
			get {return FPdfWriter;}
		}

		/// <summary>
		/// Page currently printing. 0 means the global bookmark parent of all the sheets.
		/// </summary>
		public int CurrentPage {get {return FCurrentPage;}}

		/// <summary>
		/// Page currently printing, relative to the active sheet.
		/// </summary>
		public int CurrentPageInSheet {get {return FCurrentPageInSheet;}}

		/// <summary>
		/// Bookmark that we are about to include.
		/// </summary>
		public TBookmark Bookmark {get {return FBookmark;}}
	}

    /// <summary>
    /// This event will happend each time a PDF bookmark is automatically added by FlexCel. You can use it to customize the bookmark, for example change the font color or style.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
	public delegate void GetBookmarkInformationEventHandler(object sender, GetBookmarkInformationArgs e);
	#endregion
    #endregion

}
