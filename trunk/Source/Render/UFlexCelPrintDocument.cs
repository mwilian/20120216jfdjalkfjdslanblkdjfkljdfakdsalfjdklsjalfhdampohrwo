using System;
using System.Drawing.Printing;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.ComponentModel;
using System.Drawing.Text;
using FlexCel.Core;

namespace FlexCel.Render
{
	/// <summary>
	/// A PrintDocument descendant that can be used to print and preview an Excel document.
	/// </summary>
	public class FlexCelPrintDocument: PrintDocument
	{
		#region Privates
		private FlexCelRender FRenderer=null;
		int CurrentSheetPage;
		int CurrentWorkbookPage;
        int CurrentPrintArea;
		int TotalSheetPages, TotalWorkbookPages;
		int PagesToPrint;
		RectangleF[] PaintClipRect;
		TXlsCellRange PagePrintRange;
		bool FAntiAliasedText;

		TXlsCellRange FPrintRange;
		TXlsCellRange[] MyPrintRange;

		private bool FAllVisibleSheets;
		private bool FResetPageNumberOnEachSheet;

		private int SaveActiveSheet;

		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new FlexCelPrintDocument instance.
		/// </summary>
		public FlexCelPrintDocument()
		{
			FRenderer= new FlexCelRender();
			FPrintRange=new TXlsCellRange(0,0,0,0);
		}

		/// <summary>
		/// Creates a new FlexCelPrintDocument and assigns it to an ExcelFile.
		/// </summary>
		/// <param name="aWorkbook">ExcelFile containing the data this PrintDocument will print.</param>
		public FlexCelPrintDocument(ExcelFile aWorkbook): this()
		{
			Workbook=aWorkbook;
		}

		#endregion

		#region Properties
		/// <summary>
		/// The ExcelFile to print.
		/// </summary>
		[Browsable(false)]
		public ExcelFile Workbook {get {return FRenderer.Workbook;} set {FRenderer.Workbook=value;}}

		/// <summary>
		/// Select which kind of objects should not be printed or exported to pdf.
		/// </summary>
		[Category("Behavior"),
		Description("Select which kind of objects should not be printed or exported to pdf."),
		DefaultValue(THidePrintObjects.None)]
		public THidePrintObjects HidePrintObjects {get {return FRenderer.HidePrintObjects;} set {FRenderer.HidePrintObjects=value;}}

		/// <summary>
		/// When true, the text will be antialiased. Depending on the device display might look better, but it consumes more processor time.
		/// </summary>
		[Category("Behavior"),
		Description("When true, the text will be antialiased. Display might look better but it consumes more processor time."),
		DefaultValue(false)]
		public bool AntiAliasedText {get {return FAntiAliasedText;} set {FAntiAliasedText=value;}}

		/// <summary>
		/// First column to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated.
		/// </summary>
		[Category("PrintRange"),
		Description("First column to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated."),
		DefaultValue(0)]
		public int PrintRangeLeft {get {return FPrintRange.Left;} set {FPrintRange.Left=value;}}

		/// <summary>
		/// First row to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated. 
		/// </summary>
		[Category("PrintRange"),
		Description("First row to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated."),
		DefaultValue(0)]
		public int PrintRangeTop {get {return FPrintRange.Top;} set {FPrintRange.Top=value;}}
        
		/// <summary>
		/// Last column to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated.
		/// </summary>
		[Category("PrintRange"),
		Description("Last column to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated."),
		DefaultValue(0)]
		public int PrintRangeRight {get {return FPrintRange.Right;} set {FPrintRange.Right=value;}}
        
		/// <summary>
		/// Last row to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated.
		/// </summary>
		[Category("PrintRange"),
		Description("Last row to print (1 based). if this or any other PrintRange property is 0, the range will be automatically calculated."),
		DefaultValue(0)]
		public int PrintRangeBottom {get {return FPrintRange.Bottom;} set {FPrintRange.Bottom=value;}}

		/// <summary>
		/// If true, All visible sheets on the workbook will be printed. See <see cref="ResetPageNumberOnEachSheet"/> for behavior of the page number when printing multiple sheets.
		/// </summary>
		public bool AllVisibleSheets {get {return FAllVisibleSheets;} set{FAllVisibleSheets = value;}}

		/// <summary>
		/// This property only makes sense when <see cref="AllVisibleSheets"/> is true. On that case, if this property is true each sheet of
		/// the workbook will have the page number reset. For example if the xls file has 2 sheets and each has 3 pages: 
		/// When ResetPageNumberOnEachSheet = true then footers will look like "Page 1 of 3". If false, they will look like "Page 5 of 6"
		/// </summary>
		public bool ResetPageNumberOnEachSheet {get {return FResetPageNumberOnEachSheet;} set{FResetPageNumberOnEachSheet = value;}}


		/// <summary>
		/// Returns the next page that we are going to print.
		/// </summary>
        public int CurrentPage { get { return CurrentWorkbookPage; } }

		/// <summary>
		/// Returns the next page we are going to print, on the current sheet.
		/// When not printing more than one sheet, it is equivalent to <see cref="CurrentPage"/>
		/// </summary>
        public int CurrentPageInSheet { get { return CurrentSheetPage; } }


		#endregion

		#region Utilities
		private void IncCurrentPage()
		{
			CurrentWorkbookPage++;
			CurrentSheetPage++;
			if (CurrentSheetPage > TotalSheetPages) 
			{
				CurrentSheetPage = 1;
                CurrentPrintArea = 0;
				while (Workbook.ActiveSheet < Workbook.SheetCount)
				{
					Workbook.ActiveSheet++;
					if (Workbook.SheetVisible == TXlsSheetVisible.Visible) break;
				}
			}
		}

		private void GotoFirstVisibleSheet()
		{
			Workbook.ActiveSheet = 1;
			while (Workbook.ActiveSheet < Workbook.SheetCount && Workbook.SheetVisible != TXlsSheetVisible.Visible)
			{
				Workbook.ActiveSheet++;
			}
		}

		/// <summary>
		/// Returns the number of pages to print on the whole workbook or sheet (depending on the <see cref="AllVisibleSheets"/> value).
		/// </summary>
		/// <returns></returns>
		private int GetPagesToPrint(PrintPageEventArgs e, IFlxGraphics aCanvas, RectangleF VisibleClipInches100)
		{
			int TotalPages = 0;
			int SaveActiveSheet = Workbook.ActiveSheet;
			try
			{
				for (int sheet = 1; sheet <= Workbook.SheetCount; sheet++)
				{
					Workbook.ActiveSheet = sheet;
					if (Workbook.SheetVisible != TXlsSheetVisible.Visible) continue;
					TotalPages+= TotalPagesInSheet(e, aCanvas, VisibleClipInches100);
				}

			}
			finally
			{
				Workbook.ActiveSheet = SaveActiveSheet;
			}
			return TotalPages;
		}

		private int TotalPagesInSheet(PrintPageEventArgs e, IFlxGraphics aCanvas, RectangleF VisibleClipInches100)
		{
			TXlsCellRange[] MyPrintRange = FRenderer.InternalCalcPrintArea(FPrintRange);
			TXlsCellRange PagePrintRange;
			RectangleF[] PaintClipRect;

			int TotalPages;
			bool Landscape = (Workbook.PrintOptions & TPrintOptions.Orientation)==0;
			FRenderer.InitializePrint(aCanvas, ConvertToUnits(e.PageBounds, Landscape != e.PageSettings.Landscape), 
                ConvertToMargins(e.Graphics, e.PageBounds, VisibleClipInches100, e.PageSettings.Landscape != Landscape), 
                MyPrintRange, out PaintClipRect, out TotalPages, out PagePrintRange);

			return TotalPages;

		}
		#endregion

        #region Events
		/// <summary>
		/// On .NET there is no way to get the actual physical margins of a printer without a p/invoke.
		/// As FlexCel doesn't contain p/invokes (in order to be portable to any platform), it will try to guess the best value
		/// for them by averaging the printer visible clip area. If you are not getting good results, you can
		/// let FlexCel know the real printer margins on this event. Normally, you would call GetDeviceCaps for this.
		/// </summary>
		/// <remarks>
		/// See the printpreview demo to see how this event can be used.
		/// </remarks>
		[Category("Printer Information"),
		Description("On .NET there is no way to get the actual physical margins of a printerwithout a p/invoke."
		+" As FlexCel doesn't contain p/invokes (in order to be portable to any platform), it will try to guess the best value"
		+" for them by averaging the printer visible clip area. If you are not getting good results, you can"
		+" let FlexCel know the real printer margins on this event. Normally, you would call GetDeviceCaps for this.")]      
		public event PrintHardMarginsEventHandler GetPrinterHardMargins;

		/// <summary>
		/// Replace this event when creating a custom descendant of FlexCelPrintDocument.
		/// </summary>
		/// <param name="e"></param>
		protected virtual void OnGetPrinterHardMargins(PrintHardMarginsEventArgs e)
		{
			if (GetPrinterHardMargins!=null) GetPrinterHardMargins(this, e);
		}

        /// <summary>
        /// Raises the BeginPrint event. It is called after the Print method is called and before the first page of the document prints.
        /// </summary>
        /// <param name="e">A <see cref="PrintEventArgs"/> that contains the event data.</param>
        /// <remarks>
        /// This method is overridden to print an Excel document. 
        /// Before doing its own initialization, it will call <see cref="PrintDocument.OnBeginPrint"/> so you can do your own.
        /// </remarks>
        protected override void OnBeginPrint(PrintEventArgs e)
        {
			base.OnBeginPrint(e);
			SaveActiveSheet = Workbook.ActiveSheet;
            CurrentSheetPage = 1;
            CurrentPrintArea = 0;
			CurrentWorkbookPage = 1;
			TotalSheetPages = 0;
			if (AllVisibleSheets)
			{
				GotoFirstVisibleSheet();
			}
            PagesToPrint=0;
			FRenderer.CreateFontCache(); 
			Workbook.Recalc(false);
		}

		/// <summary>
		/// Raises the EndPrint event. It is called after the Print method has finished.
		/// </summary>
		/// <param name="e">A <see cref="PrintEventArgs"/> that contains the event data.</param>
		/// <remarks>
		/// This method is overridden to print an Excel document. 
		/// </remarks>
		protected override void OnEndPrint(PrintEventArgs e)
		{
			FRenderer.DisposeFontCache();
			Workbook.ActiveSheet = SaveActiveSheet;
			base.OnEndPrint (e);
		}

		/// <summary>
		/// Fires before printing each sheet, so you can print and set up your own stuff before FlexCel renders the page.
		/// 
		/// </summary>
		[Category("Misc"),
		Description("Fires before printing each sheet, so you can print and set up your own stuff before FlexCel renders the page.")]      
		public event PrintPageEventHandler BeforePrintPage;
		
		/// <summary>
		/// Replace this event when creating a custom descendant of FlexCelPrintDocument.
		/// </summary>
		/// <param name="e"></param>
		protected virtual void OnBeforePrintPage(PrintPageEventArgs e)
		{
			if (BeforePrintPage!=null) BeforePrintPage(this, e);
		}


        //Margins are ALWAYS on inches/100, regardless of the PageUnit setting.
        private static RectangleF ConvertToUnits(Rectangle PageBounds, bool Rotate90)
        {
            float XConv= 1;// 75F/100F; //Display units are on INCHES/100  (Bad docs? they say inches/75, but they are not) 
            float YConv= 1; //75F/100F;
			if (Rotate90)
			{
				return new RectangleF(PageBounds.Top*YConv, PageBounds.Left*XConv, PageBounds.Height*YConv, PageBounds.Width*XConv);
			}
            return new RectangleF(PageBounds.Left*XConv, PageBounds.Top*YConv, PageBounds.Width*XConv, PageBounds.Height*YConv);
        }

		private RectangleF ConvertToMargins(Graphics aGraphics, Rectangle PageBounds, RectangleF VisibleRecInches100, bool Rotate90)
		{
			int w = Rotate90? PageBounds.Height: PageBounds.Width;
			int h = Rotate90? PageBounds.Width: PageBounds.Height;

			float vw = Rotate90? VisibleRecInches100.Height: VisibleRecInches100.Width;
			float vh = Rotate90? VisibleRecInches100.Width: VisibleRecInches100.Height;

			PrintHardMarginsEventArgs ex = new PrintHardMarginsEventArgs(aGraphics, 
				(w - vw)/2f, 
				(h - vh)/2f
				);

			if (ex.XMargin>0 || ex.YMargin>0) //Avoid calling when preview.
				OnGetPrinterHardMargins(ex);
		
			return new RectangleF (ex.XMargin, ex.YMargin,
					vw,
					vh);
		}

		/// <summary>
		/// Overrides the standard method to change landscape or portrait settings when <see cref="AllVisibleSheets"></see> is true. You can override this method 
        /// on a descendant class to change the behavior./>
		/// </summary>
		/// <param name="e">Arguments of the event.</param>
        protected override void OnQueryPageSettings(QueryPageSettingsEventArgs e)
		{
			base.OnQueryPageSettings (e);
			if (AllVisibleSheets)
			{
				e.PageSettings.Landscape = (Workbook.PrintOptions & TPrintOptions.Orientation)==0;
			}
		}


        /// <summary>
        /// Raises the <see cref="PrintDocument.PrintPage"/> event. It is called before a page prints.
        /// </summary>
        /// <remarks>
        /// This method is overridden to print an Excel document. 
        /// After doing its own work, it will call <see cref="PrintDocument.PrintPage"/> so you can draw your own thing over it.
        /// </remarks>
        /// <param name="e">A <see cref="PrintPageEventArgs"/> that contains the event data.</param>
        protected override void OnPrintPage(PrintPageEventArgs e)
        {
			// There is a bug in Vs2005 (but not in 2003) where the VisibleClipBounds property does not change with the PageUnit, and it remains the same as when the graphics was created.
			// See http://connect.microsoft.com/VisualStudio/feedback/ViewFeedback.aspx?FeedbackID=230594  (VisibleClipBounds does not produce correct results for different PageUnit)
			// So, we are going to store the VisibleClipBounds here. It looks like if we do not call e.Graphics.Save, things work fine.


			// Other bug in the framework (and this is also in .NET 1.1)
			//This code alone in an OnPrintPage event:
			/*
			e.Graphics.PageUnit=GraphicsUnit.Inch;
			RectangleF r = e.Graphics.VisibleClipBounds;
			e.Graphics.PageUnit=GraphicsUnit.Point;

			Font fnt = new Font("arial", 12, FontStyle.Underline);
			e.Graphics.DrawString("hello", fnt, Brushes.Tomato, 100, 100);
			
			will not print the underline, even if printing to a virtual printer. 
			(preview will work fine). If instead of GraphicsUnit.Inch we use GraphicsUnit.Display,
			Underline gets drawn in the wrong place!
			
			Using GraphicsUnit.Point in both places seems to fix it. 
			*/

			RectangleF VisibleClipInches100;
			GraphicsUnit SaveUnits = e.Graphics.PageUnit;
            try
            {
                e.Graphics.PageUnit = GraphicsUnit.Point;  //NEEDS TO BE Point, since this is what we will use later. Other unit will confuse GDI+.
                VisibleClipInches100 = e.Graphics.VisibleClipBounds;
                VisibleClipInches100 = new RectangleF(VisibleClipInches100.X / 72F * 100, VisibleClipInches100.Y / 72F * 100,
                    VisibleClipInches100.Width / 72F * 100, VisibleClipInches100.Height / 72F * 100);
            }
            finally
            {
                e.Graphics.PageUnit = SaveUnits;
            }



			if (FRenderer == null || FRenderer.Workbook==null) return;
            System.Drawing.Drawing2D.GraphicsState InitialState = e.Graphics.Save();
            try
            {
                IFlxGraphics aCanvas= new GdiPlusGraphics(e.Graphics);

				//We could change the interpolation mode for images here.
				//e.Graphics.InterpolationMode = InterpolationMode.HighQualityBilinear;
				//e.Graphics.SmoothingMode = SmoothingMode.HighQuality;

                //e.Graphics.PageUnit=GraphicsUnit.Display;  //Display is not reliable as unit as it might change depending on the device. We will be using Point.
				e.Graphics.PageUnit=GraphicsUnit.Point;

				if (AntiAliasedText) e.Graphics.TextRenderingHint=TextRenderingHint.AntiAliasGridFit;

                if (CurrentSheetPage == 1)
                {
                    FRenderer.SetCanvas(aCanvas);
                    MyPrintRange = FRenderer.InternalCalcPrintArea(FPrintRange);

                    if (CurrentWorkbookPage == 1 && AllVisibleSheets)
                    {
                        TotalWorkbookPages = GetPagesToPrint(e, aCanvas, VisibleClipInches100);
                    }
                    FRenderer.InitializePrint(aCanvas, ConvertToUnits(e.PageBounds, false),
                        ConvertToMargins(e.Graphics, e.PageBounds, VisibleClipInches100, false), MyPrintRange, out PaintClipRect, out TotalSheetPages, out PagePrintRange);

                    if (!AllVisibleSheets)
                    {
                        TotalWorkbookPages = TotalSheetPages;
                    }

                    PagesToPrint = TotalWorkbookPages;

                    if (CurrentWorkbookPage == 1) //skip the first non printing pages.
                    {
                        if (PrinterSettings.FromPage > 0 && PrinterSettings.ToPage >= PrinterSettings.FromPage)
                        {
                            PagesToPrint = Math.Min(PrinterSettings.ToPage, TotalWorkbookPages);
                            for (int i = 1; i < PrinterSettings.FromPage; i++)
                            {
                                if (CurrentWorkbookPage > TotalWorkbookPages) return;
                                OnBeforePrintPage(e);
                                FRenderer.GenericPrint(aCanvas, ConvertToUnits(e.PageBounds, false), MyPrintRange, 0, PaintClipRect, 0, false, PagePrintRange, ref CurrentPrintArea);
                                IncCurrentPage();
                            }
                        }
                    }
                }

                if (CurrentWorkbookPage <= PagesToPrint)
                {
                    int CurrentPage = ResetPageNumberOnEachSheet ? CurrentSheetPage : CurrentWorkbookPage;
                    int TotalPages = ResetPageNumberOnEachSheet ? TotalSheetPages : TotalWorkbookPages;
                    OnBeforePrintPage(e);
                    FRenderer.GenericPrint(aCanvas, ConvertToUnits(e.PageBounds, false), MyPrintRange, CurrentPage, PaintClipRect, TotalPages, true, PagePrintRange, ref CurrentPrintArea);
                    IncCurrentPage();
                }
                e.HasMorePages = CurrentWorkbookPage <= PagesToPrint;

            }
            finally
            {
                e.Graphics.Restore(InitialState);
            }
            base.OnPrintPage (e);
        }
        #endregion

	}

	#region Event args
	/// <summary>
	/// Arguments passed on <see cref="FlexCel.Render.FlexCelPrintDocument.GetPrinterHardMargins"/>. 
	/// </summary>
	public class PrintHardMarginsEventArgs: EventArgs
	{
		private float FXMargin;
		private float FYMargin;
		private Graphics FGraphics;

		/// <summary>
		/// Creates a new Argument.
		/// </summary>
		public PrintHardMarginsEventArgs(Graphics aGraphics, float aXMargin, float aYMargin)
		{
			FXMargin = aXMargin;
			FYMargin = aYMargin;
			FGraphics = aGraphics;
		}

		/// <summary>
		/// The X printer physical margin in inches / 100.
		/// </summary>
		public float XMargin {get{return FXMargin;} set{FXMargin=value;}}

		/// <summary>
		/// The Y printer physical margin in inches / 100.
		/// </summary>
		public float YMargin {get{return FYMargin;} set{FYMargin=value;}}

		/// <summary>
		/// Graphics context you can use to call GetHdc()
		/// </summary>
		public Graphics Graphics {get{return FGraphics;} set{FGraphics=value;}}
	}

    /// <summary>
    /// Delegate for getting the printer physical margins.
    /// </summary>
    public delegate void PrintHardMarginsEventHandler(object sender, PrintHardMarginsEventArgs e);

	#endregion
}
