using System;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Diagnostics;

using FlexCel.Core;
using FlexCel.Render;
using System.Globalization;
using System.Collections.Generic;

#if (WPF)
using RectangleF = System.Windows.Rect;
#else
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Security.Permissions;
using System.Security;
#endif

namespace FlexCel.Winforms
{
	/// <summary>
	/// A Simple replacement for PrintPreviewControl that allows you to preview even if the user has no
	/// printers installed. View the demo on Custom Preview to see how it is used.
	/// </summary>
	public class FlexCelPreview : System.Windows.Forms.UserControl
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

        /// <summary>
        /// Creates a new instance of FlexCelPreview.
        /// </summary>
		public FlexCelPreview()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			AutoScroll = true;
			PageXSeparation = 10;
			PageYSeparation = 10;
			BackColor = Color.Gray;
			FZoom = 1;
			FCacheSize = 64;
			PageInfo = new TPageInfoList(this);
			NumberSep = DefaultFont.Height;

			this.MouseWheel += new MouseEventHandler(OnMouseWheel); 
#if(FRAMEWORK20)
			this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
#else
			this.SetStyle(ControlStyles.DoubleBuffer | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
#endif
			this.UpdateStyles();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			try
			{
				if( disposing )
				{
					if( components != null )
						components.Dispose();
			
					PageInfo.Dispose();
				}
			}
			finally
			{
				base.Dispose(disposing);
			}
		}

		#region Component Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.SuspendLayout();
            // 
            // FlexCelPreview
            // 
            this.Name = "FlexCelPreview";
            this.ResumeLayout(false);

		}
		#endregion

		#region Privates
		private FlexCelImgExport FDocument;
		private TPageInfoList PageInfo;
		private int FPageXSeparation;
        private bool FCenteredPreview;
		private int FPageYSeparation;
		private float FZoom;
		private int FTotalPages;
		private int FCacheSize;
		private int SavedStartPage = -1;
		private FlexCelPreview FThumbnailSmall;
		private FlexCelPreview FThumbnailLarge;
		private int FThumbnailPos;
		private int NumberSep;

		private SmoothingMode FSmoothingMode;
		private InterpolationMode FInterpolationMode;

		internal TImgExportInfo FirstPageExportInfo;

		private const int ShadowSize = 3;
		#endregion

		#region Utilities
		private int GetPageNo(int height)
		{
			if (height < 0 || FirstPageExportInfo == null) return 1;
			int AcumHeight = 0;
			int AcumPages = 0;
			int SheetCount = FirstPageExportInfo.SheetCount;
			for (int i = 1; i <= SheetCount; i++)
			{
				TOneImgExportInfo SheetInfo = FirstPageExportInfo.Sheet(i);
				if (SheetInfo == null) continue;
				int SheetPages = SheetInfo.TotalPages;
				int OneSheetHeight = ((int)Math.Round(SheetInfo.PageBounds.Height*Zoom) + RealYSep);
				int SheetHeight = SheetPages * OneSheetHeight;
				if (AcumHeight + SheetHeight> height)
				{
					//Found the correct sheet.
					int RealHeight = height - AcumHeight;
			        return 1 + AcumPages + RealHeight / OneSheetHeight;
				}

				AcumHeight += SheetHeight;
				AcumPages += SheetPages;
			}
			return FirstPageExportInfo.TotalPages;
		}

		private int GetAcumPageHeight(int pageNo)
		{
			if (pageNo < 1 || FirstPageExportInfo == null) return 0;
			int AcumHeight = 0;
			int AcumPages = 0;
			int SheetCount = FirstPageExportInfo.SheetCount;
			for (int i = 1; i <= SheetCount; i++)
			{
				TOneImgExportInfo SheetInfo = FirstPageExportInfo.Sheet(i);
				if (SheetInfo == null) continue;
				int SheetPages = SheetInfo.TotalPages;
				int OneSheetHeight = ((int)Math.Round(SheetInfo.PageBounds.Height*Zoom) + RealYSep);
				int SheetHeight = SheetPages * OneSheetHeight;
				if (pageNo <= AcumPages + SheetPages)
				{
					return AcumHeight + (pageNo - AcumPages)*OneSheetHeight;
				}

				AcumHeight += SheetHeight;
				AcumPages += SheetPages;
			}

			return AcumHeight; 
		}

		private int MaxPageWidth
		{
			get
			{
				if(FirstPageExportInfo == null) return 0;
				int Result = 0;
				int SheetCount = FirstPageExportInfo.SheetCount;
				for (int i = 1; i <= SheetCount; i++)
				{
					TOneImgExportInfo SheetInfo = FirstPageExportInfo.Sheet(i);
					if (SheetInfo == null) continue;
					int w = (int)Math.Round(SheetInfo.PageBounds.Width);
					if (w > Result) Result = w;
				}
				return (int)Math.Round(Result *Zoom);
			}
		}

		private void GetVirtualPageCoords(int Page, out int Height, out int Width)
		{
			Height = 0; Width = 0;
			if (Page < 1 || FirstPageExportInfo == null) return;
			int AcumPages = 0;
			int SheetCount = FirstPageExportInfo.SheetCount;
			for (int i = 1; i <= SheetCount; i++)
			{
				TOneImgExportInfo SheetInfo = FirstPageExportInfo.Sheet(i);
				if (SheetInfo == null) continue;
				int SheetPages = SheetInfo.TotalPages;
				if (Page <= AcumPages + SheetPages)
				{
					Height = ((int)Math.Round(SheetInfo.PageBounds.Height*Zoom));
					Width  = ((int)Math.Round(SheetInfo.PageBounds.Width*Zoom));
					return;
				}
				AcumPages += SheetPages;
			}
		}


		private int GetVisiblePages()
		{
			//not 100% fool prof, but it does not need to be exact anyway.
			int h,w;
			GetVirtualPageCoords(ThumbnailPos, out h, out w);
			if (h <= 0) return 0;
			return ClientSize.Height / (h + RealYSep);
            
		}

		private int RealYSep 
		{
			get 
			{
				if (FThumbnailLarge != null) return FPageYSeparation+ NumberSep*2;
				return FPageYSeparation;
			}
		}

		private int StartPageNoMargin
		{
			get
			{
				int Result =  GetPageNo(-AutoScrollPosition.Y- PageYSeparation);
				if (Result<1) return 1;
				return Result;
			}
		}

		private void SetOnlyStartPage(int value)
		{
			AutoScrollPosition = new Point (-AutoScrollPosition.X, GetAcumPageHeight(value-1));
		}
		#endregion

		#region Events
		/// <summary>
		/// Fires when the starting page changes.
		/// </summary>
		[Category("Property Changed"),
		Description("Fires when the starting page changes.")]      
		public event EventHandler StartPageChanged;

		/// <summary>
		/// Replace this event when creating a custom descendant of FlexCelPreview.
		/// </summary>
		/// <param name="e"></param>
		protected virtual void OnStartPageChanged(EventArgs e)
		{
			if (StartPageChanged!=null) StartPageChanged(this, e);
		}

		/// <summary>
		/// Fires when the Zoom changes. (for example, the user uses ctrl+MouseWeel).
		/// </summary>
		[Category("Property Changed"),
		Description("Fires when the Zoom changes. (for example, the user uses ctrl+MouseWeel).")]      
		public event EventHandler ZoomChanged;

		/// <summary>
		/// Replace this event when creating a custom descendant of FlexCelPreview.
		/// </summary>
		/// <param name="e"></param>
		protected virtual void OnZoomChanged(EventArgs e)
		{
			if (ZoomChanged!=null) ZoomChanged(this, e);
		}


		#endregion

		#region Properties
		/// <summary>
		/// Document to be Previewed.
		/// </summary>
		[Category("Behavior"),
		Description("Document to be Previewed.")]      
		public FlexCelImgExport Document {get{return FDocument;} set{FDocument=value;}}

		/// <summary>
        /// Separation (in display units) between a page an the next. Note that if <see cref="CenteredPreview"/> is true and the preview window is bigger than the page being displayed, this value has no effect.
		/// </summary>
		[Category("Behavior"),
		Description("Separation (in display units) between a page an the next.")]      
		public int PageXSeparation {get{return FPageXSeparation;} set{if (value>0) FPageXSeparation=value;}}

        /// <summary>
        /// When true, the preview will be drawn at the middle of the window, instead of at the left. If true, then <see cref="PageXSeparation"/> is the minimum margin that the preview will have.
        /// </summary>
        [Category("Behavior"),
        Description("When true, the preview will be drawn at the middle of the window, instead of at the left.")]
        public bool CenteredPreview { get { return FCenteredPreview; } set { FCenteredPreview = value; Invalidate(); } }

		/// <summary>
		/// Separation (in display units) between a page an the next.
		/// </summary>
		[Category("Behavior"),
		Description("Separation (in display units) between a page an the next.")]      
		public int PageYSeparation {get{return FPageYSeparation;} set{if (value>0) FPageYSeparation=value;}}

		/// <summary>
		/// Page the preview is showing.
		/// </summary>
		[Browsable(false),
		Description("Page the preview is showing.")]      
		public int StartPage
		{
			get 
			{
				int Result = GetPageNo(-AutoScrollPosition.Y);
				if (Result<1) return 1;
				return Result;
			}

			set
			{
				SetOnlyStartPage(value);
				UpdateStartPage();
			}
		}

		/// <summary>
		/// Number of pages displaying.
		/// </summary>
		[Browsable(false),
		Description("Number of pages displaying.")]      
		public int TotalPages
		{
			get
			{
				return FTotalPages;
			}
		}

		/// <summary>
		/// Zoom preview.
		/// </summary>
		[Category("Behavior"),
		Description("Zoom Preview.")]      
		public float Zoom {get{return FZoom;} 
			set
			{
				if (value<0.1f) value = 0.1f;
				if (value>4f) value = 4f;
				if (value != FZoom) 
				{
					double ox = (double)AutoScrollPosition.X/(double)FZoom;
					double oy = (double)AutoScrollPosition.Y/(double)FZoom;
					FZoom=value;
					ResizeCanvas(new Point(-(int)Math.Round(ox*FZoom), -(int)Math.Round(oy*FZoom)));
					PageInfo.ClearBitmaps();
					Invalidate();
					OnZoomChanged(new EventArgs());
				}
			}
		}

		/// <summary>
		/// The cache size in number of pages stored at 100% zoom. For larger zoom
		/// actual number of pages is decreased by (Zoom*Zoom)
		/// </summary>
		[Category("Behavior"),
		Description("The cache size in number of pages stored at 100% zoom.")]      
		public int CacheSize 
		{
			get{return FCacheSize;} 
			set
			{
				if (FCacheSize<0) return;
				FCacheSize=value;
			}
		}


		/// <summary>
		/// When using this component on Thumbnail mode, set this property to another FlexCelPreview component that
		/// will hold the small Thumbnail images.
		/// </summary>
		[Category("Behavior"),
		Description("When using this component on Thumbnail mode, set this property to another FlexCelPreview component that will hold the small Thumbnail images.")]      
		public FlexCelPreview ThumbnailSmall {get{return FThumbnailSmall;} 
			set
			{
				if (FThumbnailSmall==this) return;
				FThumbnailSmall=value;
				if (FThumbnailSmall !=null)
				{
					FThumbnailLarge = null;
					FThumbnailSmall.FThumbnailLarge = this;
					FThumbnailSmall.FThumbnailSmall = null;
				}
			}
		}

		/// <summary>
		/// When using this component on Thumbnail mode, set this property to another FlexCelPreview component that
		/// will hold the large Thumbnail images.
		/// </summary>
		[Category("Behavior"),
		Description("When using this component on Thumbnail mode, set this property to another FlexCelPreview component that will hold the large Thumbnail images.")]      
		public FlexCelPreview ThumbnailLarge {get{return FThumbnailLarge;} 
			set
			{
				if (FThumbnailLarge==this) return;
				FThumbnailLarge=value;
				Zoom = 0.10f;
				if (FThumbnailLarge !=null)
				{
					FThumbnailSmall = null;
					FThumbnailLarge.FThumbnailSmall = this;
					FThumbnailLarge.FThumbnailLarge = null;
				}
			}
		}

		/// <summary>
		/// This affects how the images are rendered on the screen. Some modes will look a little blurred but with better quality.
		/// Consult the .NET framework documentation on SmoothingMode for more information
		/// </summary>
		[Category("Appearance"),
		Description("This affects how the images are rendered on the screen. Some modes will look a little blurred but with better quality.")]      
		public SmoothingMode SmoothingMode {get{return FSmoothingMode;} set{FSmoothingMode=value;}}

		/// <summary>
		/// This affects how the images are rendered on the screen. Some modes will look a little blurred but with better quality.
		/// Consult the .NET framework documentation on SmoothingMode for more information
		/// </summary>
		[Category("Appearance"),
		Description("This affects how the images are rendered on the screen. Some modes will look a little blurred but with better quality.")]      
		public InterpolationMode InterpolationMode {get{return FInterpolationMode;} set{FInterpolationMode=value;}}



		#endregion

		#region Public
		/// <summary>
		/// Invalidates the preview and forces the control to reload from the document. 
		/// When the control is a Thumbnail you cannot Invalidate it, this will be done automatically when you invalidate the main view.
		/// </summary>
		public void InvalidatePreview()
		{
			if (FThumbnailLarge !=null) throw new InvalidOperationException("InvalidatePreview should be called on the main display component. This will also invalidate the Thumbnails");
			ReloadDocument();
			if (FThumbnailSmall !=null)
			{
				FThumbnailSmall.ReloadDocument();
				FThumbnailSmall.Invalidate();
			}
			Invalidate();
		}

		#endregion

        #region Implementation
        private void ScrollIntoView(int Page)
		{
			int PagePos = FPageYSeparation + GetAcumPageHeight((Page-1));
			int YPos = -AutoScrollPosition.Y;

			if (PagePos<YPos) 
			{
				StartPage = Page;
				return;
			}

			int VirtualPageHeight, VirtualPageWidth;
			GetVirtualPageCoords(Page, out VirtualPageHeight, out VirtualPageWidth);
			if (PagePos+ VirtualPageHeight+ RealYSep*2>= YPos + ClientSize.Height)
			{
				AutoScrollPosition = new Point (-AutoScrollPosition.X, (PagePos+ VirtualPageHeight+ RealYSep*2 - ClientSize.Height));
				UpdateStartPage();
				return;
			}
		}

		internal int ThumbnailPos {get{return FThumbnailPos;} 
			set
			{
				if (value>TotalPages) value = TotalPages;
				if (value <1) value = 1;
				ScrollIntoView(value);
				FThumbnailPos=value;
				Invalidate();
			}
		}

		private void OnMouseWheel(object sender, MouseEventArgs e)
		{
			if (ThumbnailLarge != null) return;
			if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
			{
				Zoom += e.Delta/120F/10F;
			}
		}

		private void UpdateMainView()
		{
			if (FThumbnailLarge !=null)
				FThumbnailLarge.SetOnlyStartPage(FThumbnailPos);
		}

        /// <summary>
        /// Overrides the standard mousedown event to handle it.
        /// </summary>
        /// <param name="e"></param>
		protected override void OnMouseDown(MouseEventArgs e)
		{
			base.OnMouseDown (e);
			if (FThumbnailLarge !=null)
			{
				ThumbnailPos = GetPageNo(e.Y-AutoScrollPosition.Y-FPageYSeparation);		
				UpdateMainView();
			}
		}

		private bool HandleThumbKey(Keys e)
		{
			switch (e)
			{
				case Keys.Down:
				case Keys.Right:
					ThumbnailPos++;
					UpdateMainView();
					return true;

				case Keys.Up:
				case Keys.Left:
					ThumbnailPos--;
					UpdateMainView();
					return true;

				case Keys.Down | Keys.Control:
				case Keys.PageDown:
					int i1 =GetVisiblePages();
					if (i1<1) i1=1;
					ThumbnailPos+= i1;
					UpdateMainView();
					return true;

				case Keys.Up | Keys.Control:
				case Keys.PageUp:
					int i2 =GetVisiblePages();
					if (i2<1) i2=1;
					ThumbnailPos-= i2;
					UpdateMainView();
					return true;

				case Keys.PageUp | Keys.Control:
				case Keys.Home | Keys.Control:
				case Keys.Home:
					ThumbnailPos =1;
					UpdateMainView();
					return true;

				case Keys.PageDown | Keys.Control:
				case Keys.End | Keys.Control:
				case Keys.End:
					ThumbnailPos = TotalPages;
					UpdateMainView();
					return true;
			}
			return false;
		}

		private bool HandleMainKey(Keys e)
		{
			switch (e)
			{
				case Keys.Down:
					AutoScrollPosition = new Point (-AutoScrollPosition.X, -AutoScrollPosition.Y+10);
					Invalidate();
					return true;

				case Keys.Up:
					AutoScrollPosition = new Point (-AutoScrollPosition.X, -AutoScrollPosition.Y-10);
					Invalidate();
					return true;

				case Keys.Down | Keys.Control:
				case Keys.PageDown:
					AutoScrollPosition = new Point (-AutoScrollPosition.X, -AutoScrollPosition.Y+ClientSize.Height);
					Invalidate();
					return true;

				case Keys.Up | Keys.Control:
				case Keys.PageUp:
					AutoScrollPosition = new Point (-AutoScrollPosition.X, -AutoScrollPosition.Y-ClientSize.Height);
					Invalidate();
					return true;

				case Keys.Left | Keys.Control:
					AutoScrollPosition = new Point (-AutoScrollPosition.X-ClientSize.Width, -AutoScrollPosition.Y);
					Invalidate();
					return true;

				case Keys.Left:
					AutoScrollPosition = new Point (-AutoScrollPosition.X-10, -AutoScrollPosition.Y);
					Invalidate();
					return true;

				case Keys.Right | Keys.Control:
					AutoScrollPosition = new Point (-AutoScrollPosition.X+ClientSize.Width, -AutoScrollPosition.Y);
					Invalidate();
					return true;

				case Keys.Right:
					AutoScrollPosition = new Point (-AutoScrollPosition.X+10, -AutoScrollPosition.Y);
					Invalidate();
					return true;

				case Keys.PageUp | Keys.Control:
				case Keys.Home | Keys.Control:
				case Keys.Home:
					StartPage = 1;
					return true;

				case Keys.PageDown | Keys.Control:
				case Keys.End | Keys.Control:
				case Keys.End:
					AutoScrollPosition = new Point (-AutoScrollPosition.X, AutoScrollMinSize.Height);
					Invalidate();
					return true;
			}
			return false;
		}

        #if (!FULLYMANAGED)
        /// <summary>
        /// Overrides the ProcessCmdKey event.
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="keyData"></param>
        /// <returns></returns>
#if (FRAMEWORK40)
        [SecuritySafeCritical]
#else
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        [SecurityPermission(SecurityAction.InheritanceDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
#endif
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
		{
			if (FThumbnailLarge !=null)
			{
				if (HandleThumbKey(keyData)) return true;
			}
			else
			{
				if (HandleMainKey(keyData)) return true;
			}

			return base.ProcessCmdKey (ref msg, keyData);
        }
        #endif

        /// <summary>
        /// Overrides the OnfontChanged event.
        /// </summary>
        /// <param name="e"></param>
		protected override void OnFontChanged(EventArgs e)
		{
			base.OnFontChanged (e);
			NumberSep = DefaultFont.Height;
		}

        /// <summary>
        /// Overrides de OnResize method.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnResize(EventArgs e)
        {
            if (CenteredPreview) Invalidate();
            base.OnResize(e);
        }

		private void ReloadDocument()
		{
			if (Document == null) return;
			PageInfo.Clear();
			if (FThumbnailLarge !=null)  //Reload document is always called first for the big preview and later for the thumbs.
				FirstPageExportInfo = TImgExportInfo.Clone(FThumbnailLarge.FirstPageExportInfo);
			else
				FirstPageExportInfo = Document.GetFirstPageExportInfo();

			FTotalPages = FirstPageExportInfo.TotalPages;
			SavedStartPage = 0;

			ResizeCanvas(new Point(0,0));
			OnZoomChanged(new EventArgs());
		}

		private void ResizeCanvas(Point NewPosition)
		{
            AutoScrollMinSize = new Size(MaxPageWidth + PageXSeparation * 2,
                FPageYSeparation + GetAcumPageHeight(TotalPages));

			AutoScrollPosition = NewPosition;
			UpdateStartPage();
		}

		private void DrawPageNumber(Graphics g, int Page, int YPos)
		{
            string PageStr = (Page + 1).ToString(CultureInfo.CurrentCulture);
			SizeF sz = g.MeasureString(PageStr, DefaultFont);
			Brush bgBrush = null;
			Brush fgBrush = Brushes.White;

			if (Page == ThumbnailPos-1)
			{
				bgBrush = Brushes.Navy;
				fgBrush = Brushes.White;
			}

			int VirtualPageHeight, VirtualPageWidth;
			GetVirtualPageCoords(Page + 1, out VirtualPageHeight, out VirtualPageWidth);
            int CurrentPosX = GetCurrentPosX(VirtualPageWidth);
            if (bgBrush != null) g.FillRectangle(bgBrush, CurrentPosX + 2, YPos + VirtualPageHeight + 4, VirtualPageWidth - 4, NumberSep + 2);
            g.DrawString(PageStr, DefaultFont, fgBrush, CurrentPosX + VirtualPageWidth / 2 - sz.Width / 2, YPos + VirtualPageHeight + 5);        
		}

        private int GetCurrentPosX(int PageWidth)
        {
            int Result = PageXSeparation + AutoScrollPosition.X;
            if (CenteredPreview && ClientSize.Width > PageWidth + 2 * PageXSeparation)
            {
                Result += (ClientSize.Width - (PageWidth + 2 * PageXSeparation)) / 2;
            }

            return Result;
        }

        /// <summary>
        /// Overrides the OnPaint event.
        /// </summary>
        /// <param name="pe"></param>
		protected override void OnPaint(PaintEventArgs pe)
		{		
			base.OnPaint(pe);
			pe.Graphics.InterpolationMode = InterpolationMode;
			pe.Graphics.SmoothingMode = SmoothingMode;

			if (Document == null)
			{
				pe.Graphics.DrawString("No Document assigned.", DefaultFont, Brushes.Black, 
					0 + AutoScrollPosition.X, 0 + AutoScrollPosition.Y);
				return;
			}


			int CurrentPage = StartPageNoMargin-1;
			int CurrentPosY = FPageYSeparation + GetAcumPageHeight(CurrentPage) + AutoScrollPosition.Y;
			while (CurrentPage < TotalPages && CurrentPosY<= pe.ClipRectangle.Bottom+2)
			{
				int VirtualPageHeight, VirtualPageWidth;
				GetVirtualPageCoords(CurrentPage + 1, out VirtualPageHeight, out VirtualPageWidth);

                int CurrentPosX = GetCurrentPosX(VirtualPageWidth);
                pe.Graphics.DrawRectangle(Pens.Black, CurrentPosX - 1, CurrentPosY - 1, VirtualPageWidth + 2, VirtualPageHeight + 2);
                pe.Graphics.FillRectangle(Brushes.Black, CurrentPosX + VirtualPageWidth + 1, CurrentPosY + ShadowSize, ShadowSize, VirtualPageHeight);
                pe.Graphics.FillRectangle(Brushes.Black, CurrentPosX + ShadowSize, CurrentPosY + VirtualPageHeight + 1, VirtualPageWidth, ShadowSize);
	
				pe.Graphics.DrawImage(PageInfo.GetImage(CurrentPage), CurrentPosX , CurrentPosY, VirtualPageWidth, VirtualPageHeight);

				if (ThumbnailLarge != null)
					DrawPageNumber(pe.Graphics, CurrentPage, CurrentPosY);

				CurrentPage++;
                CurrentPosY += VirtualPageHeight + RealYSep;
			}

			if (ThumbnailLarge != null && TotalPages >0)
			{
				int ThumbPosY = FPageYSeparation + GetAcumPageHeight(ThumbnailPos-1) + AutoScrollPosition.Y;
				int VirtualPageHeight, VirtualPageWidth;
				GetVirtualPageCoords(ThumbnailPos, out VirtualPageHeight, out VirtualPageWidth);

                int CurrentPosX = GetCurrentPosX(VirtualPageWidth);
                pe.Graphics.DrawRectangle(Pens.Navy, CurrentPosX - 2, ThumbPosY - 2, VirtualPageWidth + 4, VirtualPageHeight + 4);
                pe.Graphics.DrawRectangle(Pens.Navy, CurrentPosX - 3, ThumbPosY - 3, VirtualPageWidth + 6, VirtualPageHeight + 6);
			}

			UpdateStartPage();
		}

		private void UpdateThumbs()
		{
			if (ThumbnailSmall != null)
			{
				ThumbnailSmall.ThumbnailPos = StartPage;
			}
		}

		private void UpdateStartPage()
		{
			if (StartPage != SavedStartPage)
			{
				UpdateThumbs();
				SavedStartPage = StartPage;
				OnStartPageChanged(new EventArgs());
			}
		}

        #endregion
    }

    #region Utility Classes

    /// <summary>
	/// Holds a list of page information items.
	/// </summary>
	internal class TPageInfoList: List<TPageInfo>, IDisposable
	{
		FlexCelPreview Owner;

		#region Cache
		TPageInfo CacheFirst;
		TPageInfo CacheLast;
		int CacheCount;
		#endregion

		public TPageInfoList (FlexCelPreview aOwner)
		{
			Owner = aOwner;
		}

		public new void Clear()
		{
			Dispose();
			base.Clear();
			CacheCount = 0;
			CacheFirst = null;
			CacheLast = null;
		}

		private bool GetNextImage(ref TImgExportInfo ExportInfo, Bitmap bmp)
		{
			if (bmp == null) return Owner.Document.ExportNext(null, ref ExportInfo);
			using (Graphics gr = Graphics.FromImage(bmp))
			{
				gr.Clear(Color.White);
				gr.InterpolationMode = Owner.InterpolationMode;
				gr.SmoothingMode = Owner.SmoothingMode;
				return Owner.Document.ExportNext(gr, ref ExportInfo);
			}
		}

		internal Bitmap GetImage(int index)
		{
			TImgExportInfo ei = null;
			if (Count>0)
				ei = TImgExportInfo.Clone(this[Count-1].ExportInfo);
			else ei = TImgExportInfo.Clone(Owner.FirstPageExportInfo);

			for (int i = Count; i <= index; i++)
			{
				GetNextImage(ref ei, null);
				Add (new TPageInfo(TImgExportInfo.Clone(ei), null));
			}

			TPageInfo Pi = this[index];
			if (Pi.Bmp == null)
			{
				RectangleF Page = Pi.ExportInfo.ActiveSheet.PageBounds;
				Pi.Bmp = BitmapConstructor.CreateBitmap((int)(Page.Width * Owner.Zoom), (int)(Page.Height * Owner.Zoom));
				Pi.Bmp.SetResolution(96*Owner.Zoom, 96*Owner.Zoom);
				TImgExportInfo ExportInfo = null;
				if (index>0) ExportInfo = TImgExportInfo.Clone(this[index-1].ExportInfo);
				else
					ExportInfo = TImgExportInfo.Clone(Owner.FirstPageExportInfo);

				GetNextImage(ref ExportInfo, Pi.Bmp);
				CacheCount++;
			}

			//Bring the item to last position on the cache.
			if (Pi.Next!=null) 
			{
				Pi.Next.Prev = Pi.Prev;
				if (Pi.Prev!=null) Pi.Prev.Next = Pi.Next;
				else CacheFirst = Pi.Next;
			}
			if (CacheFirst ==null) CacheFirst = Pi;
			
			if (CacheLast != Pi)
			{
				Pi.Prev = CacheLast;
				if (CacheLast!=null) CacheLast.Next = Pi;
				CacheLast = Pi;
				Pi.Next = null;
			}

			//If we have more bitmaps on the cache that what is allowed, delete the one at the first position.
			if (CacheFirst !=null && CacheCount > 1 + Owner.CacheSize / (Owner.Zoom * Owner.Zoom))
			{
				if (CacheFirst.Next != null)
				{
					CacheFirst.Bmp.Dispose();
					CacheFirst.Bmp = null;
					CacheCount--;
					
					TPageInfo Next = CacheFirst.Next;
					Next.Prev = null;
					CacheFirst.Next = null;
					CacheFirst = Next;
				}

			}

			Debug.Assert(CacheFirst == null || CacheFirst.Prev ==null,"Error in cache");
			Debug.Assert(CacheLast == null || CacheLast.Next ==null,"Error in cache");
			return Pi.Bmp;

		}

		internal void ClearBitmaps()
		{
			TPageInfo Pi = CacheFirst;
			while (Pi !=null)
			{
				Pi.Bmp.Dispose();
				Pi.Bmp=null;
				Pi.Prev=null;
				TPageInfo TmpPi = Pi;
				Pi = TmpPi.Next;
				TmpPi.Next = null;
			}
			CacheCount=0;
			CacheFirst = null;
			CacheLast = null;

		}

		#region IDisposable Members

		public void Dispose()
		{
			for (int i = Count-1; i>=0; i--)
			{
				this[i].Dispose();
			}
            GC.SuppressFinalize(this);
        }

		#endregion
	}

	/// <summary>
	/// Holds information about how to render a page, and optionally a cache with a bitmap.
	/// </summary>
	internal class TPageInfo: IDisposable
	{
		internal Bitmap Bmp;
		internal TImgExportInfo ExportInfo;
		
		internal TPageInfo Prev;
		internal TPageInfo Next;

		internal TPageInfo(TImgExportInfo aExportInfo, Bitmap aBmp)
		{
			Bmp = aBmp;
			ExportInfo = aExportInfo;
			Prev = null;
			Next = null;
		}

		#region IDisposable Members

		public void Dispose()
		{
			if (Bmp !=null) Bmp.Dispose();
            GC.SuppressFinalize(this);
        }
		#endregion

	}




	#endregion
}
