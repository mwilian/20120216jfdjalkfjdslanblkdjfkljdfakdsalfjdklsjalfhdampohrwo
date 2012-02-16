using System;

#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	#else
	using System.Drawing;
	using System.Drawing.Imaging;
	#endif
#endif
using System.ComponentModel;
using System.IO;
using FlexCel.Core;
using System.Security.Permissions;
using System.Security;

namespace FlexCel.Render
{
    /// <summary>
    /// A component for exporting an Excel file to an image. It can return an image object, or the
    /// actual bytes of an specific file format. (like gif, tiff or png)
    /// </summary>
	public class FlexCelImgExport: Component
	{
		#region Privates
		private FlexCelRender FRenderer=null;

		bool FAllowOverwritingFiles;

		TXlsCellRange FPrintRange;

		TPaperDimensions FPageSize;

		private float FResolution;

		private bool FAllVisibleSheets;
		private bool FResetPageNumberOnEachSheet;

		private int LastInitSheet;
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new FlexCelImgExport instance.
		/// </summary>
		public FlexCelImgExport()
		{
			FRenderer = new FlexCelRender();
			FPrintRange = new TXlsCellRange(0,0,0,0);
			PageSize = null;
			Resolution = 96;
		}

		/// <summary>
		/// Creates a new FlexCelImgExport and assigns it to an ExcelFile.
		/// </summary>
		/// <param name="aWorkbook">ExcelFile containing the data this component will export.</param>
		public FlexCelImgExport(ExcelFile aWorkbook): this()
		{
			Workbook=aWorkbook;
		}

		/// <summary>
		/// Creates a new FlexCelImgExport and assigns it to an ExcelFile, setting AllowOverwritingFiles to the desired value.
		/// </summary>
		/// <param name="aWorkbook">ExcelFile containing the data this component will export.</param>
		/// <param name="aAllowOverwritingFiles">When true, existing files will be overwrited.</param>
		public FlexCelImgExport(ExcelFile aWorkbook, bool aAllowOverwritingFiles): this()
		{
			Workbook=aWorkbook;
			FAllowOverwritingFiles = aAllowOverwritingFiles;
		}

		#endregion

		#region Properties
		/// <summary>
		/// The ExcelFile to print.
		/// </summary>
		[Browsable(false)]
		public ExcelFile Workbook {get {return FRenderer.Workbook;} set {FRenderer.Workbook=value;}}

		/// <summary>
		/// Select which kind of objects should not be printed or exported to the image.
		/// </summary>
		[Category("Behavior"),
		Description("Select which kind of objects should not be printed or exported to the image."),
		DefaultValue(THidePrintObjects.None)]
		public THidePrintObjects HidePrintObjects {get {return FRenderer.HidePrintObjects;} set {FRenderer.HidePrintObjects=value;}}

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
		/// Image page size. Set it to null to use the paper size on the xls file.
		/// </summary>
		public TPaperDimensions PageSize 
		{
			get {return FPageSize;} 
			set 
			{
				FPageSize=value; 
			}
		}


		/// <summary>
		/// Determines if FlexCel will automatically delete existing image files or not.
		/// </summary>
		[Category("Behavior"),
		Description("Determines if FlexCel will automatically delete existing image files or not."),
		DefaultValue(false)]
		public bool AllowOverwritingFiles {get {return FAllowOverwritingFiles;} set {FAllowOverwritingFiles=value;}}
		
		/// <summary>
		/// "The default resolution on pixels per inch for the rendered images. For the screen, this is 96."
		/// </summary>
		[Category("Behavior"),
		Description("The default resolution on pixels per inch for the rendered images. For the screen, this is 96."),
		DefaultValue(96)]
		public float Resolution {get{return FResolution;} set{FResolution=value;}}

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

		#endregion

		#region Events
		/// <summary>
		/// Fires before drawing the image, allowing to modify it or to modify the XlsFile associated.
		/// </summary>
		[Category("Behavior"),
		Description("Fires before drawing the image, allowing to modify it or to modify the XlsFile associated.")]      
		public event PaintEventHandler BeforePaint;

		/// <summary>
		/// Replace this event when creating a custom descendant of FlexCelImgExport.
		/// </summary>
		/// <param name="e"></param>
		protected virtual void OnBeforePaint(ImgPaintEventArgs e)
		{
			if (BeforePaint != null) BeforePaint(this, e);
		}

		/// <summary>
		/// Fires after the image has been drawn, allowing to modify it.
		/// </summary>
		[Category("Behavior"),
		Description("Fires after the image has been drawn, allowing to modify it.")]      
		public event PaintEventHandler AfterPaint;

		/// <summary>
		/// Replace this event when creating a custom descendant of FlexCelImgExport.
		/// </summary>
		/// <param name="e"></param>
		protected virtual void OnAfterPaint(ImgPaintEventArgs e)
		{
			if (AfterPaint != null) AfterPaint(this, e);
		}

		#endregion

		#region Export
		/// <summary>
		/// Exports the associated xls workbook to a graphics stream. You need to provide a 
		/// Graphics object with the correct dimensions. (To get the needed dimensions, use <see cref="GetRealPageSize()"/>
		/// </summary>
		/// <param name="imgData">Graphics where the image will be stored. Set it to null to skip the page.</param>
		/// <param name="exportInfo"> Information needed to export, cached for speed. The first time you call this method (or when you change xls.ActiveSheet), make exportInfo=null</param>
        public bool ExportNext(Graphics imgData, ref TImgExportInfo exportInfo)
        {
            FRenderer.CreateFontCache();
            try
            {
                Bitmap bmp = null;

                try
                {
                    if (imgData == null)
                    {
                        bmp = BitmapConstructor.CreateBitmap(1, 1);
                        imgData = Graphics.FromImage(bmp);
                        imgData.PageUnit = GraphicsUnit.Point;
                    }
                    IFlxGraphics aCanvas = new GdiPlusGraphics(imgData);

                    GraphicsUnit OriginalUnits = imgData.PageUnit;
                    try
                    {
                        imgData.PageUnit = GraphicsUnit.Point;
                        FRenderer.SetCanvas(aCanvas);
                        try
                        {
                            if (exportInfo == null) exportInfo = GetExportInfo(aCanvas);

                            exportInfo.IncCurrentPage();
                            if (exportInfo.CurrentPage > exportInfo.TotalPages) return false;

                            int SaveActiveSheet = Workbook.ActiveSheet;
                            try
                            {
                                Workbook.ActiveSheet = exportInfo.CurrentSheet;
                                int CurrentLogicalPage = -1;
                                if (ResetPageNumberOnEachSheet)
                                {
                                    CurrentLogicalPage = exportInfo.ActiveSheet.FCurrentPage;
                                }
                                else
                                {
                                    CurrentLogicalPage = exportInfo.CurrentPage;
                                }


                                TOneImgExportInfo OneResult = exportInfo.ActiveSheet;
                                if (LastInitSheet != exportInfo.CurrentSheet)
                                {
                                    TXlsCellRange ra; int p; RectangleF[] r;
                                    FRenderer.InitializePrint(aCanvas, OneResult.PageBounds, OneResult.PageBounds, OneResult.PrintRanges, out r, out p, out ra);
                                    LastInitSheet = exportInfo.CurrentSheet;
                                }

                                if (bmp == null)
                                {
                                    OnBeforePaint(new ImgPaintEventArgs(imgData, CalcPageBounds(exportInfo.ActiveSheet.PageBounds), exportInfo.CurrentPage, exportInfo.ActiveSheet.CurrentPage, exportInfo.TotalPages));
                                }

                                FRenderer.GenericPrint(aCanvas, OneResult.PageBounds, OneResult.PrintRanges, CurrentLogicalPage,
                                    OneResult.PaintClipRect, exportInfo.TotalLogicalPages(ResetPageNumberOnEachSheet), bmp == null, 
                                    OneResult.PagePrintRange, ref OneResult.FCurrentPrintArea);

                                aCanvas.ResetClip();

                                if (bmp == null)
                                {
                                    OnAfterPaint(new ImgPaintEventArgs(imgData, CalcPageBounds(exportInfo.ActiveSheet.PageBounds), exportInfo.CurrentPage, exportInfo.ActiveSheet.CurrentPage, exportInfo.TotalPages));
                                }
                            }
                            finally
                            {
                                Workbook.ActiveSheet = SaveActiveSheet;
                            }
                        }
                        finally
                        {
                            FRenderer.SetCanvas(null);
                        }
                    }
                    finally
                    {
                        imgData.PageUnit = OriginalUnits;
                    }
                }
                finally
                {
                    if (bmp != null)
                    {
                        bmp.Dispose();
                        imgData.Dispose();
                    }
                }
            }
            finally
            {
                FRenderer.DisposeFontCache();
            }
            return true;
        }

		private static RectangleF CalcPageBounds(RectangleF PageBounds)
		{
			return new RectangleF(PageBounds.Left * 72f / 100f, PageBounds.Top * 72f / 100f, PageBounds.Width * 72f / 100f, PageBounds.Height * 72f / 100f); 
		}

		/// <summary>
		/// Exports the associated xls workbook to a stream.
		/// </summary>
		/// <param name="imgStream">Stream where the image will be exported.</param>
		/// <param name="format">Pixel depth for the created image.</param>
		/// <param name="imgFormat">Format for the saved image</param>
		/// <param name="exportInfo"> Information needed to export, cached for speed. The first time you call this method (or when you change xls.ActiveSheet), make exportInfo=null</param>
		public bool ExportNext(Stream imgStream, PixelFormat format, ImageFormat imgFormat, ref TImgExportInfo exportInfo)
		{
			bool Result = false;
			using (Bitmap bmp = CreateBitmap(Resolution, ref exportInfo, format))
			{
				using (Graphics g = Graphics.FromImage(bmp))
				{
					Result = ExportNext(g, ref exportInfo);
				}
				if (Result) bmp.Save(imgStream, imgFormat);
			}
			return Result;
		}

		/// <summary>
		/// Exports the associated xls workbook to a file.
		/// </summary>
		/// <param name="fileName">File to export.</param>
		/// <param name="format">Format for the created image.</param>
		/// <param name="imgFormat">Format for the saved image</param>
		/// <param name="exportInfo"> Information needed to export, cached for speed. The first time you call this method (or when you change xls.ActiveSheet), make exportInfo=null</param>
		public bool ExportNext(string fileName, PixelFormat format, ImageFormat imgFormat, ref TImgExportInfo exportInfo)
		{
			bool Result = false;
			if (exportInfo!=null && exportInfo.CurrentPage>=exportInfo.TotalPages) return false;
			//if (Workbook.RowCount<=0) return false; A sheet might have only images.
			try
			{
				FileMode fm=FileMode.CreateNew;
				if (AllowOverwritingFiles) fm=FileMode.Create;
				using (FileStream f= new FileStream(fileName, fm, FileAccess.Write))
				{
					Result = ExportNext(f, format, imgFormat, ref exportInfo);
				}
			}
            catch (IOException)
            {
                //Don't delete the file in an io exception. It might be because allowoverwritefiles was false, and the file existed.
                throw;
            }
            catch 
			{
				File.Delete(fileName);
				throw;
			}

			return Result;
		}
	

		/// <summary>
		/// Return the pages to print. This is a costly operation, so cache the results.
		/// </summary>
		/// <returns></returns>
		public int TotalPages()
		{
			TImgExportInfo ei = null;
			if (!ExportNext(null, ref ei)) return 0;
			return ei.TotalPages;
		}

		/// <summary>
		/// Returns the page dimensions for a sheet. You can use it to create a bitmap to export the data.
		/// </summary>
		/// <returns></returns>
		public TPaperDimensions GetRealPageSize(int ActiveSheet)
		{
			TPaperDimensions pd = FPageSize;
			if (FPageSize==null)
			{
				int SaveActiveSheet = Workbook.ActiveSheet;
				try
				{
					Workbook.ActiveSheet = ActiveSheet;

					pd = Workbook.PrintPaperDimensions;
					if ((Workbook.PrintOptions & TPrintOptions.Orientation)==0)
					{
						float w = pd.Width;
						pd.Width = pd.Height;
						pd.Height = w;
					}
				}
				finally
				{
					Workbook.ActiveSheet = SaveActiveSheet;
				}
			}
			return new TPaperDimensions(pd.PaperName, pd.Width, pd.Height);
		}

		/// <summary>
		/// Returns the page dimensions for the active sheet. You can use it to create a bitmap to export the data.
		/// </summary>
		/// <returns></returns>
		public TPaperDimensions GetRealPageSize()
		{
			return GetRealPageSize(Workbook.ActiveSheet);
		}

		/// <summary>
		/// Returns information needed for exporting multiple pages on one sheet. You normally 
		/// don't need to use this method, but you can use it to speed up multiple displays.
		/// </summary>
		/// <returns></returns>
		public TImgExportInfo GetFirstPageExportInfo()
		{
			FRenderer.CreateFontCache();
            try
            {
                using (Bitmap bmp = BitmapConstructor.CreateBitmap(1, 1))
                {
                    using (Graphics imgData = Graphics.FromImage(bmp))
                    {
                        IFlxGraphics aCanvas = new GdiPlusGraphics(imgData);

                        imgData.PageUnit = GraphicsUnit.Point;
                        FRenderer.SetCanvas(aCanvas);
                        try
                        {
                            return GetExportInfo(aCanvas);
                        }
                        finally
                        {
                            FRenderer.SetCanvas(null);
                        }

                    }
                }
            }
            finally
            {
                FRenderer.DisposeFontCache();
            }
		}


		private TImgExportInfo GetExportInfo(IFlxGraphics aCanvas)
		{
			TImgExportInfo Result = new TImgExportInfo();

			if (AllVisibleSheets)
			{
				int FirstVisibleSheet = -1;
				int SaveActiveSheet = Workbook.ActiveSheet;
				try
				{
					Result.Sheets = new TOneImgExportInfo[Workbook.SheetCount];
					for (int sheet = 1; sheet <=Workbook.SheetCount; sheet ++)
					{
						Workbook.ActiveSheet = sheet;
						if (Workbook.SheetVisible != TXlsSheetVisible.Visible) continue;
						if (FirstVisibleSheet < 0) FirstVisibleSheet  = sheet;

						TOneImgExportInfo OneResult = new TOneImgExportInfo();
				
						OneResult.FPrintRanges = FRenderer.InternalCalcPrintArea(FPrintRange);
						TPaperDimensions pd = GetRealPageSize();
						OneResult.FPageBounds = new RectangleF(0, 0, pd.Width, pd.Height);
						
						FRenderer.InitializePrint(aCanvas, OneResult.PageBounds, OneResult.PageBounds, OneResult.PrintRanges, out OneResult.FPaintClipRect, out OneResult.FTotalPages, out OneResult.FPagePrintRange);        
						Result.Sheets[sheet-1] = OneResult;
					}
				}
				finally
				{
					Workbook.ActiveSheet = SaveActiveSheet;
				}

				Result.CurrentSheet = FirstVisibleSheet;
			}
			else
			{
				TOneImgExportInfo OneResult = new TOneImgExportInfo();
				OneResult.FPrintRanges = FRenderer.InternalCalcPrintArea(FPrintRange);
				TPaperDimensions pd = GetRealPageSize();
				OneResult.FPageBounds = new RectangleF(0, 0, pd.Width, pd.Height);

				FRenderer.InitializePrint(aCanvas, OneResult.PageBounds, OneResult.PageBounds, OneResult.PrintRanges, out OneResult.FPaintClipRect, out OneResult.FTotalPages, out OneResult.FPagePrintRange);        
				Result.Sheets = new TOneImgExportInfo[1];
				Result.Sheets[0] = OneResult;
				Result.CurrentSheet = Workbook.ActiveSheet;

			}
			Result.ResetCurrentPage();
			LastInitSheet = 0;
			return Result;
		}

		#endregion

        #region Export to Images
        #region Public interface
        /// <summary>
        /// Saves the current Excel file as an image file. 
        /// </summary>
        /// <param name="fileName">File where the image will be saved.</param>
        /// <param name="export">Image format.</param>
        /// <param name="ColorDepth">Color depth for the image, if applicable. Some formats (like fax, that is monochrome) do not allow different color depths.</param>
#if (!FULLYMANAGED)
#if (FRAMEWORK40)
        [SecuritySafeCritical]
#else
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
#endif
#endif
        public void SaveAsImage(string fileName, ImageExportType export, ImageColorDepth ColorDepth)
        {
            try
            {
                FileMode fm = FileMode.CreateNew;
                if (AllowOverwritingFiles) fm = FileMode.Create;
                using (FileStream f = new FileStream(fileName, fm, FileAccess.Write))
                {
                    SaveAsImage(f, export, ColorDepth);
                }
            }
            catch (IOException)
            {
                //Don't delete the file in an io exception. It might be because allowoverwritefiles was false, and the file existed.
                throw;
            }
            catch
            {
                File.Delete(fileName);
                throw;
            }
        }

        /// <summary>
        /// Saves the current Excel file on an image stream. 
        /// </summary>
        /// <param name="fileStream">Stream where the image will be saved.</param>
        /// <param name="export">Image format.</param>
        /// <param name="ColorDepth">Color depth for the image, if applicable. Some formats (like fax, that is monochrome) do not allow different color depots.</param>
#if (!FULLYMANAGED)
#if (FRAMEWORK40)
        [SecuritySafeCritical]
#else
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
#endif
#endif
        public void SaveAsImage(Stream fileStream, ImageExportType export, ImageColorDepth ColorDepth)
        {
            if (Workbook == null) FlxMessages.ThrowException(FlxErr.ErrWorkbookNull);
            Workbook.Recalc(false);

            switch (export)
            {
                case ImageExportType.Tiff:
                    CreateMultiPageTiff(fileStream, ColorDepth, export);
                    break;

                case ImageExportType.Fax:
                    CreateMultiPageTiff(fileStream, ColorDepth, export);
                    break;

                case ImageExportType.Fax4:
                    CreateMultiPageTiff(fileStream, ColorDepth, export);
                    break;

                case ImageExportType.Gif:
                    CreateImg(fileStream, ImageFormat.Gif, ImageColorDepth.Color256);
                    break;

                case ImageExportType.Png:
                    CreateImg(fileStream, ImageFormat.Png, ColorDepth);
                    break;

                case ImageExportType.Jpeg:
                    CreateImg(fileStream, ImageFormat.Jpeg, ColorDepth);
                    break;

                default: FlxMessages.ThrowException(FlxErr.ErrInvalidImageFormat);
                    break;
            }
        }
        #endregion

        #region Multipage Tiff
        private static ImageCodecInfo GetTiffEncoder()
        {
            foreach (ImageCodecInfo ImageEncoder in ImageCodecInfo.GetImageEncoders())
                if (ImageEncoder.MimeType == "image/tiff")
                    return ImageEncoder;

            FlxMessages.ThrowException(FlxErr.ErrTiffEncoderNotFound);
            return null; //just to compile.
        }

        private Bitmap CreateBitmap(float Resolution, ref TImgExportInfo ExportInfo, PixelFormat PxFormat)
        {
			if (ExportInfo == null)
			{
				ExportInfo = GetFirstPageExportInfo();
			}

			TPaperDimensions pd = GetRealPageSize(ExportInfo.NextSheet);
            Bitmap Result =
                BitmapConstructor.CreateBitmap((int)Math.Ceiling(pd.Width / 96F * Resolution),
                (int)Math.Ceiling(pd.Height / 96F * Resolution), PxFormat);
            Result.SetResolution(Resolution, Resolution);
            return Result;
        }

#if (!FULLYMANAGED)
#if (FRAMEWORK40)
        [SecuritySafeCritical]
#else
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
#endif
#endif
        private void CreateMultiPageTiff(Stream OutStream, ImageColorDepth ColorDepth, ImageExportType ExportType)
        {
           ImageCodecInfo info = GetTiffEncoder();

            int ParamCount = 1;
            if (ExportType == ImageExportType.Fax || ExportType == ImageExportType.Fax4) ParamCount++;

            EncoderParameters ep = new EncoderParameters(ParamCount);
            ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.MultiFrame);

            bool IsFax = false;

            switch (ExportType)
            {
                case ImageExportType.Fax:
                    ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT3);
                    IsFax = true;
                    break;
                case ImageExportType.Fax4:
                    ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);
                    IsFax = true;
                    break;
            }

            bool Is1bpp = IsFax || ColorDepth == ImageColorDepth.BlackAndWhite;

            TImgExportInfo ExportInfo = null;

            PixelFormat RgbPixFormat = IsFax || ColorDepth != ImageColorDepth.TrueColor ? PixelFormat.Format32bppPArgb : PixelFormat.Format24bppRgb;
            PixelFormat PixFormat = PixelFormat.Format1bppIndexed;
			if (!IsFax) 
			{
				switch (ColorDepth)
				{
					case ImageColorDepth.TrueColor: PixFormat = RgbPixFormat; break;
					case ImageColorDepth.Color256: PixFormat = PixelFormat.Format8bppIndexed; break;
				}
			}

            using (Bitmap OutImg = CreateBitmap(Resolution, ref ExportInfo, PixFormat))
            {

                //First image is handled differently.
                Bitmap ActualOutImg = Is1bpp || ColorDepth != ImageColorDepth.TrueColor ? CreateBitmap(Resolution, ref ExportInfo, RgbPixFormat) : OutImg;
                try
                {
                    using (Graphics Gr = Graphics.FromImage(ActualOutImg))
                    {
                        Gr.FillRectangle(Brushes.White, 0, 0, ActualOutImg.Width, ActualOutImg.Height); //Clear the background
                        ExportNext(Gr, ref ExportInfo);
                    }

                    if (Is1bpp) FloydSteinbergDither.ConvertToBlackAndWhite(ActualOutImg, OutImg);
                    else
                        if (!IsFax && ColorDepth == ImageColorDepth.Color256)
                    {
                        OctreeQuantizer.ConvertTo256Colors(ActualOutImg, OutImg);
                    }
                }
                finally
                {
                    if (ActualOutImg != OutImg) ActualOutImg.Dispose();
                }

                OutImg.Save(OutStream, info, ep);
                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);


                //Now the rest of images.
                ExportInfo = ExportAllImagesButFirst(ep, Is1bpp, ExportInfo, RgbPixFormat, OutImg, ColorDepth);

                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.Flush);
                OutImg.SaveAdd(ep);

            }
        }

#if (!FULLYMANAGED)
#if (FRAMEWORK40)
        [SecuritySafeCritical]
#else
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
#endif
#endif
        private TImgExportInfo ExportAllImagesButFirst(EncoderParameters ep, bool Is1bpp, TImgExportInfo ExportInfo, PixelFormat RgbPixFormat, Bitmap OutImg, ImageColorDepth ColorDepth)
        {
            if (ExportInfo == null) ExportInfo = GetFirstPageExportInfo();
            for (int i = ExportInfo.CurrentPage; i < ExportInfo.TotalPages; i++)
            {
                using (Bitmap TmpImg = CreateBitmap(Resolution, ref ExportInfo, RgbPixFormat))
                {
                    using (Graphics Gr = Graphics.FromImage(TmpImg))
                    {
                        Gr.FillRectangle(Brushes.White, 0, 0, TmpImg.Width, TmpImg.Height); //Clear the background
                        ExportNext(Gr, ref ExportInfo);

                        if (Is1bpp)
                        {
                            using (Bitmap BwImg = FloydSteinbergDither.ConvertToBlackAndWhite(TmpImg))
                            {
                                OutImg.SaveAdd(BwImg, ep);
                            }
                        }
                        else
                            if (ColorDepth == ImageColorDepth.Color256)
                        {
                            using (Bitmap IndexImg = OctreeQuantizer.ConvertTo256Colors(TmpImg))
                            {
                                OutImg.SaveAdd(IndexImg, ep);
                            }
                        }
                        else
                        {
                            OutImg.SaveAdd(TmpImg, ep);
                        }
                    }
                }
            }
            return ExportInfo;
        }
        #endregion
        #region Simple Images
#if (!FULLYMANAGED)
#if (FRAMEWORK40)
        [SecuritySafeCritical]
#else
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
#endif
#endif
        private void CreateImg(Stream OutStream, ImageFormat ImgFormat, ImageColorDepth Colors)
        {
            TImgExportInfo ExportInfo = null;

            PixelFormat RgbPixFormat = Colors != ImageColorDepth.TrueColor ? PixelFormat.Format32bppPArgb : PixelFormat.Format24bppRgb;
            PixelFormat PixFormat = PixelFormat.Format1bppIndexed;
            switch (Colors)
            {
                case ImageColorDepth.TrueColor: PixFormat = RgbPixFormat;break;
                case ImageColorDepth.Color256: PixFormat = PixelFormat.Format8bppIndexed;break;
            }

            using (Bitmap OutImg = CreateBitmap(Resolution, ref ExportInfo, PixFormat))
            {
                Bitmap ActualOutImg = Colors != ImageColorDepth.TrueColor? CreateBitmap(Resolution, ref ExportInfo, RgbPixFormat): OutImg;
                try
                {
                    using (Graphics Gr = Graphics.FromImage(ActualOutImg))
                    {
                        Gr.FillRectangle(Brushes.White, 0, 0, ActualOutImg.Width, ActualOutImg.Height); //Clear the background
                        ExportNext(Gr, ref ExportInfo);
                    }

                    if (Colors == ImageColorDepth.BlackAndWhite) FloydSteinbergDither.ConvertToBlackAndWhite(ActualOutImg, OutImg);
                    else
                        if (Colors == ImageColorDepth.Color256)
                    {
                        OctreeQuantizer.ConvertTo256Colors(ActualOutImg, OutImg);
                    }
                }
                finally
                {
                    if (ActualOutImg != OutImg) ActualOutImg.Dispose();
                }

                OutImg.Save(OutStream, ImgFormat);
            }
        }
        #endregion
        #endregion
	}

	/// <summary>
	/// Holds information needed to export the pages, so it is only calculated once.
	/// </summary>
	public class TImgExportInfo
	{
		#region Private variables.
		/// <summary>
		/// This is the current sheet on the worksheet. Not in the FSheets array.
		/// </summary>
		private int FCurrentSheet;
		private int FCurrentPage;
		private TOneImgExportInfo[] FSheets;
		#endregion

		/// <summary>
		/// Returns the Image export info for one of the sheets.
		/// </summary>
		/// <param name="index">Sheet on the list (1 based).</param>
		/// <returns></returns>
		public TOneImgExportInfo Sheet(int index)
		{
			return FSheets[index - 1];
		}

		/// <summary>
		/// Return the count of the sheets on the workbook.
		/// </summary>
		public int SheetCount
		{
			get
			{
				if (FSheets == null) return 0;
				return FSheets.Length;
			}
		}

		internal TOneImgExportInfo[] Sheets{get{return FSheets;} set {FSheets = value;}}


		/// <summary>
		/// Sheet that is being printed.
		/// </summary>
		public int CurrentSheet {get {return FCurrentSheet;} set{FCurrentSheet = value;}}

		/// <summary>
		/// Sheet of the next page to print.
		/// </summary>
		public int NextSheet 
		{
			get 
			{
				if (ActiveSheet.CurrentPage < ActiveSheet.TotalPages) return CurrentSheet;
				int cs = CurrentSheet + 1;
				while (cs < FSheets.Length && FSheets[cs -1] == null) cs++;

				return cs;
			}
		}

		/// <summary>
		/// Last page printed.
		/// </summary>
		public int CurrentPage {get{return FCurrentPage;}}

		/// <summary>
		/// Total pages to print for all the sheets.
		/// </summary>
		public int TotalPages
		{
			get
			{
				if (FSheets == null) return 0;
				int Result = 0;
				for (int i = 0; i < FSheets.Length; i++)
				{
					if (FSheets[i] != null) Result += FSheets[i].TotalPages;
				}
				return Result;
			}
		}

        /// <summary>
        /// TImageInfo for the active sheet.
        /// </summary>
		public TOneImgExportInfo ActiveSheet
		{
			get
			{
				if (FSheets == null) return null;
				if (FSheets.Length == 1) return FSheets[0]; //When exporting only the active sheet FSheets will be length 1.
				return FSheets[CurrentSheet - 1];
			}
		}

		/// <summary>
		/// Returns the total count of pages if ResetPageNumberOnEachSheet is false, or the page number for the current sheet if true.
		/// </summary>
		public int TotalLogicalPages(bool ResetPageNumberOnEachSheet)
		{
			if (FSheets == null) return 0;
			if (ResetPageNumberOnEachSheet) return ActiveSheet.TotalPages;
			return TotalPages;
		}

		internal void ResetCurrentPage()
		{
			FCurrentPage = 0;
			if (Sheets != null)
			{
				foreach (TOneImgExportInfo ImgInfo in Sheets)
				{
					if (ImgInfo != null) ImgInfo.FCurrentPage = 0;
				}
			}
		}

        internal void IncCurrentPage()
        {
            FCurrentPage++;
            if (ActiveSheet != null)
            {
                ActiveSheet.FCurrentPage++;
                if (ActiveSheet.CurrentPage > ActiveSheet.TotalPages)
                {
                    do
                    {
                        CurrentSheet++;
                    }
                    while (CurrentSheet - 1 < FSheets.Length && FSheets[CurrentSheet - 1] == null);

                    if (CurrentSheet - 1 < FSheets.Length) ActiveSheet.FCurrentPage = 1;

                }
            }
        }

        /// <summary>
        /// Returns a deep copy of a TImgExportInfo. This method will work even if the source object is null.
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
		public static TImgExportInfo Clone(TImgExportInfo source)
		{
			if (source == null) return null;
			TImgExportInfo Result = (TImgExportInfo)source.MemberwiseClone();
			if (source.FSheets != null)
			{
				Result.FSheets = new TOneImgExportInfo[source.FSheets.Length];
				for (int i = 0; i < source.FSheets.Length; i++)
				{
					Result.FSheets[i] = TOneImgExportInfo.Clone(source.FSheets[i]);
				}
			}
			return Result;
		}
	}

	/// <summary>
	/// Holds information needed to export one of the workbbok sheets, so it is only calculated once.
	/// </summary>
	public class TOneImgExportInfo
	{
		internal int FCurrentPage;
		internal int FTotalPages;
        internal int FCurrentPrintArea;
		internal RectangleF[] FPaintClipRect;
		internal TXlsCellRange FPagePrintRange;

		internal RectangleF FPageBounds;
		internal TXlsCellRange[] FPrintRanges;

		/// <summary>
		/// Total pages on the ActiveSheet.
		/// </summary>
        public int TotalPages { get { return FTotalPages; } }

        /// <summary>
        /// Total pages on the ActiveSheet.
        /// </summary>
        public int CurrentPrintArea { get { return FCurrentPrintArea; } }

		/// <summary>
		/// Coordinates to print. One per every isolated range in the print area of the sheet.
		/// </summary>
		public RectangleF[] PaintClipRect {get{return FPaintClipRect;}}

		/// <summary>
		/// Range that has been printed. Note that before printing, the values here are invalid and you should use <see cref="PrintRanges"/>
		/// </summary>
        public TXlsCellRange PagePrintRange { get { return FPagePrintRange; } }

		/// <summary>
		/// Limits of the page.
		/// </summary>
        public RectangleF PageBounds { get { return FPageBounds; } }

		/// <summary>
		/// Range that will be printed.
        /// One per every isolated range in the print area of the sheet.
        /// </summary>
        public TXlsCellRange[] PrintRanges { get { return FPrintRanges; } }

        /// <summary>
        /// Range that will be printed. When the print area is composed of different non contiguous parts
        /// you should use <see cref="PrintRanges"/> to gett all the parts.
        /// </summary>
        public TXlsCellRange PrintRange { get { return FPrintRanges[0]; } }

		/// <summary>
		/// Last page printed on this sheet.
		/// </summary>
        public int CurrentPage { get { return FCurrentPage; } }

		/// <summary>
		/// Clones a TImageExportInfo even if it is null.
		/// </summary>
		/// <param name="source"></param>
		/// <returns></returns>
		public static TOneImgExportInfo Clone(TOneImgExportInfo source)
		{
			if (source == null)
				return null;

			TOneImgExportInfo Result = new TOneImgExportInfo();
			Result.FCurrentPage = source.FCurrentPage;
			Result.FTotalPages = source.FTotalPages;
            Result.FCurrentPrintArea = source.FCurrentPrintArea;
            if (source.FPaintClipRect != null)
            {
                Result.FPaintClipRect = new RectangleF[source.FPaintClipRect.Length];
                for (int i = 0; i < source.FPaintClipRect.Length; i++)
                {
                    Result.FPaintClipRect[i] = source.FPaintClipRect[i]; //struct                    
                }
            }
			Result.FPagePrintRange = (TXlsCellRange) source.FPagePrintRange.Clone();
 
			Result.FPageBounds = source.FPageBounds; //struct

            if (source.FPrintRanges != null)
            {
                Result.FPrintRanges = new TXlsCellRange[source.FPrintRanges.Length];
                for (int i = 0; i < source.FPrintRanges.Length; i++)
                {
                    Result.FPrintRanges[i] = (TXlsCellRange)source.FPrintRanges[i].Clone();                    
                }
            }
			return Result;
		}

	}

	#region Event Classes
	/// <summary>
	/// Arguments passed on Paint events.
	/// </summary>
	public class ImgPaintEventArgs: EventArgs
	{
		private readonly Graphics FGraphics;
		private readonly RectangleF FPageBounds;
		private readonly int FCurrentPage;
		private readonly int FCurrentPageInSheet;
		private readonly int FTotalPages;


		/// <summary>
		/// Creates a new Argument.
		/// </summary>
		/// <param name="aGraphics">Gets the graphics used to paint.</param>
		/// <param name="aPageBounds">Gets the rectangle in which to paint.</param>
		/// <param name="aCurrentPage">The page we are printing.</param>
		/// <param name="aCurrentPageInSheet">The page we are printing, relative to the current sheet.</param>
		/// <param name="aTotalPages">The total number of page we have available to export.</param>
		public ImgPaintEventArgs(Graphics aGraphics, RectangleF aPageBounds, int aCurrentPage, int aCurrentPageInSheet, int aTotalPages)
		{
			FGraphics = aGraphics;
			FPageBounds = aPageBounds;
			FCurrentPage = aCurrentPage;
			FCurrentPageInSheet = aCurrentPageInSheet;
			FTotalPages = aTotalPages;
		}

		/// <summary>
		/// Gets the graphics used to paint.
		/// </summary>
		public Graphics Graphics
		{
			get {return FGraphics;}
		}

		/// <summary>
		/// Gets the rectangle in which to paint.
		/// </summary>
		public RectangleF PageBounds {get {return FPageBounds;}}

		/// <summary>
		/// Gets the current page number.
		/// </summary>
		public int CurrentPage {get {return FCurrentPage;}}

		/// <summary>
		/// Gets the current page number, relative to the active sheet.
		/// </summary>
		public int CurrentPageInSheet {get {return FCurrentPageInSheet;}}

		/// <summary>
		/// Gets the total number of pages available to export.
		/// </summary>
		public int TotalPages {get {return FTotalPages;}}
	}

	/// <summary>
	/// Delegate for Paint events.
	/// </summary>
	public delegate void PaintEventHandler(object sender, ImgPaintEventArgs e);
	#endregion

    #region Utility Types
    /// <summary>
    /// Defines how you want to export the sheet.
    /// </summary>
    public enum ImageExportType
    {
        /// <summary>
        /// Export as png image. Only the first page of the Excel active sheet will be exported.
        /// </summary>
        Png = 0,

        /// <summary>
        /// Export as gif file. Only the first page of the Excel active sheet will be exported.
        /// </summary>
        Gif = 12,

        /// <summary>
        /// Export as jpg file. Only the first page of the Excel active sheet will be exported.
        /// </summary>
        Jpeg = 13,

        /// <summary>
        /// Export as multipage tiff.
        /// </summary>
        Tiff = 14,

        /// <summary>
        /// Export as a black and white multipage tiff ccitt3 compatible with standard fax.
        /// </summary>
        Fax = 15,

        /// <summary>
        /// Export as a black and white multipage tiff ccitt4 compatible with standard fax.
        /// </summary>
        Fax4 = 16
    }

    /// <summary>
    /// Number of colors for the exported images.
    /// </summary>
    public enum ImageColorDepth
    {
        /// <summary>
        /// 24 bits per pixel.
        /// </summary>
        TrueColor = 0,

        /// <summary>
        /// 8 bits per pixel, 256 colors.
        /// </summary>
        Color256 = 5,

        /// <summary>
        /// 1 bit per pixel, 2 colors (black and white).
        /// </summary>
        BlackAndWhite = 10
    }
    #endregion


}

