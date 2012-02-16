using System;
using System.ComponentModel;
using System.IO;
using System.Globalization;
using FlexCel.Core;
using System.Collections.Generic;

using System.Text;

#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	using PointF = System.Windows.Point;
	using real = System.Double;
	using System.Windows.Media;
	#else
	using real = System.Single;
	using System.Drawing;
	using System.Drawing.Drawing2D;
	using System.Drawing.Imaging;
	#endif
#endif

namespace FlexCel.Render
{
	#region FlexCelHtmlExport
    /// <summary>
    /// A component for exporting an Excel file to HTML.
    /// </summary>
	public class FlexCelHtmlExport: Component, IHtmlFontEvent
	{
		#region Privates   
		private ExcelFile FWorkbook;

		private volatile bool FCanceled;
		private volatile FlexCelHtmlExportProgress FProgress;

		THtmlVersion FHtmlVersion;
		internal THtmlFixes FHtmlFixes;
		bool FAllowOverwritingFiles;
		private THtmlFileFormat FHtmlFileFormat;

		private THidePrintObjects FHidePrintObjects;

		TXlsCellRange FPrintRange;

		private TImageProps FImageProps;
		private THtmlImageFormat FSavedImagesFormat;


		private string FClassPrefix;

		private bool FVerticalTextAsImages;
        private bool FExportNamedRanges;

		private THtmlExtraInfo FExtraInfo;

		private TGeneratedFiles FGeneratedFiles;

		private const string DefaultHeadingStyle = "background-color:#E7E7E7;text-align:center;border: 1px solid black;font-family:helvetica,arial,sans-serif;font-size:10pt";
		private real FHeadingWidth;
		private string FHeadingStyle;

		private bool FUseContentId;

		private bool FIgnoreSharingViolations;

		private string FBaseUrl;
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new FlexCelHtmlExport instance.
		/// </summary>
		public FlexCelHtmlExport()
		{
			FHtmlVersion = THtmlVersion.Html_401;
			FHtmlFileFormat = THtmlFileFormat.Html;
			FPrintRange = new TXlsCellRange(0,0,0,0);
			FProgress = new FlexCelHtmlExportProgress();
			FHtmlFixes = new THtmlFixes();
			FClassPrefix = "flx";

			FImageProps = new TImageProps();
			FImageProps.ImageResolution = 96;
			FImageProps.SmoothingMode = SmoothingMode.AntiAlias;
			FImageProps.InterpolationMode = InterpolationMode.HighQualityBicubic;
			FImageProps.AntiAliased = true;
            FImageProps.ImageBackground = ColorUtil.Empty;
			FSavedImagesFormat = THtmlImageFormat.Png;

			FHeadingStyle = DefaultHeadingStyle;
			FHeadingWidth = 50;

			FVerticalTextAsImages = true;
			FIgnoreSharingViolations = true;

			FExtraInfo = new THtmlExtraInfo();
			FGeneratedFiles = new TGeneratedFiles();
			FUseContentId = true;

			FHidePrintObjects = THidePrintObjects.HeadersAndFooters;

		}

		/// <summary>
		/// Creates a new FlexCelHtmlExport and assigns it to an ExcelFile.
		/// </summary>
		/// <param name="aWorkbook">ExcelFile containing the data this component will export.</param>
		public FlexCelHtmlExport(ExcelFile aWorkbook): this()
		{
			Workbook=aWorkbook;
		}

		/// <summary>
		/// Creates a new FlexCelHtmlExport and assigns it to an ExcelFile, setting AllowOverwritingFiles to the desired value.
		/// </summary>
		/// <param name="aWorkbook">ExcelFile containing the data this component will export.</param>
		/// <param name="aAllowOverwritingFiles">When true, existing files will be overwrited.</param>
		public FlexCelHtmlExport(ExcelFile aWorkbook, bool aAllowOverwritingFiles): this()
		{
			Workbook=aWorkbook;
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
		public bool Canceled{ get {return FCanceled;} 
			set 
			{
				if (value==true) FCanceled=true; //Don't allow to uncancel.
			}
		}

		/// <summary>
		/// Progress of the export. This variable must be accessed from other thread.
		/// </summary>
		[Browsable(false)]
		public FlexCelHtmlExportProgress Progress
		{ 
			get {return FProgress;} 
		}

		/// <summary>
		/// The ExcelFile to export.
		/// </summary>
		[Browsable(false)]
		public ExcelFile Workbook {get {return FWorkbook;} set {FWorkbook=value;}}

		/// <summary>
		/// Select which kind of objects should not be exported to html. By default we do *not* export headers and footers, since they are normally not what you want when exporting to HTML.
		/// </summary>
		[Category("Behavior"),
		Description("Select which kind of objects should not be exported to Html. By default we do *not* export headers and footers, since they are normally not what you want when exporting to HTML."),
		DefaultValue(THidePrintObjects.HeadersAndFooters)]
		public THidePrintObjects HidePrintObjects {get {return FHidePrintObjects;} set {FHidePrintObjects=value;}}

		/// <summary>
		/// Version of the HTML generated.
		/// </summary>
		[Category("Behavior"),
		Description("Version of the HTML generated."),
		DefaultValue(THtmlVersion.Html_401)]
		public THtmlVersion HtmlVersion {get {return FHtmlVersion;} set {FHtmlVersion=value;}}

		/// <summary>
		/// Format of the HTML file to be generated.
		/// </summary>
		[Category("Behavior"),
		Description("Format of the HTML file to be generated."),
		DefaultValue(THtmlFileFormat.Html)]
		public THtmlFileFormat HtmlFileFormat {get {return FHtmlFileFormat;} set{FHtmlFileFormat = value;}}


		/// <summary>
		/// Prefix to be appended to all CSS classes. For example, if you set it to "test", CSS classes will be
		/// named like ".test1234". Normally you do not need to change this property, but if you need to insert multiple
		/// Excel files in the same HTML page, you need to ensure all classes have an unique ClassPrefix.
		/// </summary>
		[Category("Behavior"),
		Description("Prefix to be appended to all CSS classes. You need to change it only if appending many Excel files in one HTML page."),
		DefaultValue("flx")]
		public string ClassPrefix {get {return FClassPrefix;} set{FClassPrefix = value;}}


		/// <summary>
		/// By default, Internet explorer does not support transparent PNGs. Normally this is not an issue, since Excel does not use 
		/// much transparency. But if you rely on transparent images and don't want to use gif images instead of png, you can set this
		/// property to true. It will add special code to the HTML file to support transparent images in IE6.<br/>
        /// <b>Note:</b> If setting this property to false, you might want to set <see cref="ImageBackground"/> to Color.White instead
        /// of ColorUtil.Empty to ensure images have no transparent background.
		/// </summary>
		[Category("Browser Fixes"),
		Description("Fix IE6 rendering bugs with transparent pngs. You do not normally need to apply this fix. See documentation for more information."),
		DefaultValue(false)]
		public bool FixIE6TransparentPngSupport {get {return FHtmlFixes.IE6TransparentPngSupport;} set{FHtmlFixes.IE6TransparentPngSupport = value;}}

		/// <summary>
		/// Outlook 2007 renders HTML worse than previous versions, since it switched to the Word 2007 rendering engine instead of
		/// Internet Explorer to show HTML emails. If you apply this fix, some code will be added to the generated HTML file to improve
		/// the display in Outlook 2007. Other browsers will not be affected and will still render the original file. Turn this option on if
		/// you plan to email the generated file as an HTML email or to edit them in Word 2007. Note that the pages will not validate with the
		/// w3c validator if this option is on.
		/// </summary>
		[Category("Browser Fixes"),
		Description("Fix Outlook 2007 rendering bugs. See documentation for more information."),
		DefaultValue(false)]
		public bool FixOutlook2007CssSupport {get {return FHtmlFixes.Outlook2007CssSupport;} set{FHtmlFixes.Outlook2007CssSupport = value;}}

		/// <summary>
		/// Some older browsers (and Word 2007) might not support the CSS white-space tag. In this case, if a line longer than a cell cannot be expanded to the right
		/// (because there is data in the next cell) it will wrap down instead of being cropped. This fix will cut the text on this cell to the displayable
		/// characters. If a letter was displayed by the half on the right, after applying this fix it will not display.
		/// This fix is automatically applied when <see cref="FixOutlook2007CssSupport"/> is selected, so there is normally no reason to apply it. You might get 
		/// a smaller file with this fix (if you have a lots of hidden text), but the display will not be as accurate as when it is off, so it is reccomended to keep it off.
		/// </summary>
		[Category("Browser Fixes"),
		Description("Cut long strings in text for browsers that will always wrap overlapping text. You do not normally need to apply this fix. See documentation for more information."),
		DefaultValue(false)]
		public bool FixIE6WordWrapSupport {get {return FHtmlFixes.WordWrapSupport;} set{FHtmlFixes.WordWrapSupport = value;}}

		/// <summary>
		/// When exporting to <b>MHTML</b>, some mail clients might have problems understanding the newer "Content-Location" header to show the images.
		/// When this property is true, we will use the older "Content-Id" header that is better supported than Content Location in the mime headers
		/// to reference the images. You are strongly encouraged to keep this property
		/// true in order to maximize the number of mail readers compatible. When Exporting to HTML (not MHTML), this property has no effect.
		/// </summary>
		[Category("Browser Fixes"),
		Description("When this is true, we will use the older \"Content-Id\" header when creating MHTML files. You are strongly encouraged to keep this property true. See documentation for more information."),
		DefaultValue(true)]
		public bool UseContentId {get {return FUseContentId;} set{FUseContentId = value;}}


		/// <summary>
		/// Width in points of the left gutter when printing row numbers and column names.
		/// </summary>
		[Category("Headings"),
		Description("Width in points of the left gutter when printing row numbers and column names."),
		DefaultValue(50)]
		public real HeadingWidth {get {return FHeadingWidth;} set{FHeadingWidth = value;}}

		/// <summary>
		/// Style definition for the gutter cells when printing row numbers or column names. This text must be 
		/// a valid CSS style definition, without including the braces ("{}").
		/// </summary>
		/// <example>If you specify "color:red;font-weight:bold" in this property, text will be red and bold.</example>
		[Category("Headings"),
		Description("Style definition for the gutter cells when printing row numbers or column names. Text must not include braces ({})"),
		DefaultValue(DefaultHeadingStyle)]
		public string HeadingStyle {get {return FHeadingStyle;} set{FHeadingStyle = value;}}

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
		/// Determines if FlexCel will automatically delete existing HTML and image files or not.
		/// </summary>
		[Category("Behavior"),
		Description("Determines if FlexCel will automatically delete existing HTML and image files or not."),
		DefaultValue(false)]
		public bool AllowOverwritingFiles {get {return FAllowOverwritingFiles;} set {FAllowOverwritingFiles=value;}}

		/// <summary>
		/// When this property is true and the component tries to write any file that is locked by other thread, it will not raise
		/// an error and just assume the other thread will write the correct image. You will normally want to have this true, so you can have many threads
		/// writing to the same file without issues. Note that when <see cref="AllowOverwritingFiles"/> is false, this property has no effect.
		/// </summary>
		[Category("Behavior"),
		Description("If this is true and there is a sharing error when saving a file, the component will ignore it and assume the file written by the other thread is correct."),
		DefaultValue(true)]
		public bool IgnoreSharingViolations {get {return FIgnoreSharingViolations;} set{FIgnoreSharingViolations = value;}}

		/// <summary>
		/// When true and text is vertical, FlexCel will replace the text with an image in order to show it correctly in HTML.
		/// When false, text will be rendered normally without rotation.
		/// </summary>
		[Category("Behavior"),
		Description("When true and text is vertical, FlexCel will replace the text with an image in order to show it correctly in HTML. When false, text will be rendered normally without rotation."),
		DefaultValue(true)]
		public bool VerticalTextAsImages {get {return FVerticalTextAsImages;} set{FVerticalTextAsImages = value;}}

        /// <summary>
        /// When true FlexCel will insert a span in the first cell of every named range with "id" = the name of the range.
        /// You can access then this with javascript.<br></br> For a fine grain control of how names are exported, you can use <see cref="NamedRangeExport"/> event.
        /// </summary>
        [Category("Behavior"),
        Description("When true FlexCel will insert a span in the first cell of every named range with id = the name of the range."),
        DefaultValue(false)]
        public bool ExportNamedRanges { get { return FExportNamedRanges; } set { FExportNamedRanges = value; } }

		/// <summary>
		/// If this property is not null, all hyperlinks stating with this value will be converted to relative links, by removing this string form them.
		/// <br/>Hyperlinks in Excel must be absolute by default, so this property is a way to get relative hyperlinks.
		/// <br/>For example, if BaseUrl is "http://www.tmssoftware.com/" and an Excel file has a link "http://www.tmssoftware.com/test.html"
		/// the link in the generated HTML file will be "test.html"
		/// </summary>
		[Category("Behavior"),
		Description("If this property is not null, all hyperlinks stating with this value will be converted to relative links, by removing this string form them."),
		DefaultValue(null)]
		public string BaseUrl {get {return FBaseUrl;} set{FBaseUrl = value;}}


		/// <summary>
		/// Resolution for the exported images. The bigger the resolution, the bigger the image size and quality. Use 96 for standard screen resolution.
		/// </summary>
		[Category("Images"),
		Description("Resolution for the exported images. The bigger the resolution, the bigger the image size and quality. Use 96 for standard screen resolution."),
		DefaultValue(96)]
		public real ImageResolution {get {return FImageProps.ImageResolution;} set{FImageProps.ImageResolution = value;}}

		/// <summary>
		/// File format in which the images will be saved. Note that Ie6 does not support transparency in PNGs by default,
		/// so if you have transparent images and you want to make you page compatible with IE6, you should save as gif or use <see cref="FixIE6TransparentPngSupport"/>
		/// </summary>
		[Category("Images"),
		Description("File format to save the images. See docs for more information about png support."),
		DefaultValue(THtmlImageFormat.Png)]
		public THtmlImageFormat SavedImagesFormat {get {return FSavedImagesFormat;} set{FSavedImagesFormat = value;}}

		/// <summary>
		/// This affects how the images, charts, etc are rendered for the image file. Some modes will look a little blurred but with better quality.
		/// Consult the .NET framework documentation on SmoothingMode for more information
		/// </summary>
		[Category("Images"),
		Description("This affects how the images, charts, etc are rendered for the image files. Some modes will look a little blurred but with better quality."),
		DefaultValue(SmoothingMode.AntiAlias)]      
		public SmoothingMode SmoothingMode {get{return FImageProps.SmoothingMode;} set{FImageProps.SmoothingMode=value;}}

		/// <summary>
		/// This affects how the images, charts, etc are rendered for the image file. Some modes will look a little blurred but with better quality.
		/// Consult the .NET framework documentation on SmoothingMode for more information
		/// </summary>
		[Category("Images"),
		Description("This affects how the images, charts, etc are rendered for the image file. Some modes will look a little blurred but with better quality."),
		DefaultValue(InterpolationMode.HighQualityBicubic)]      
		public InterpolationMode InterpolationMode {get{return FImageProps.InterpolationMode;} set{FImageProps.InterpolationMode=value;}}
		
		/// <summary>
		/// This affects how the text is rendered for example when exporting a chart. Some modes will look a little blurred but with better quality.
		/// Consult the .NET framework documentation on SmoothingMode for more information
		/// </summary>
		[Category("Images"),
		Description("This affects how the text is rendered for example when exporting a chart. Some modes will look a little blurred but with better quality."),
		DefaultValue(true)]      
		public bool AntiAliased {get{return FImageProps.AntiAliased;} set{FImageProps.AntiAliased=value;}}

        /// <summary>
        /// When this property is set to ColorUtil.Empty (the default), images will be rendered with a transparent background. 
        /// While this is the normal behavior, sometimes you might not want transparent images (for example to support Internet Explorer 6
        /// without setting <see cref="FixIE6TransparentPngSupport"/> to true), and then you could use Color.White here.
        /// </summary>
        [Category("Images"),
        Description("When this property is set to ColorUtil.Empty (the default), images will be rendered with a transparent background. Set it to Color.White to support IE6 without setting FixIE6TransparentPngSupport = true"),
        ]
        public Color ImageBackground { get { return FImageProps.ImageBackground; } set { FImageProps.ImageBackground = value; } }

		/// <summary>
		/// This property defines how the images will be named by FlexCel. You can always override the name using the <see cref="GetImageInformation"/> event.
		/// </summary>
		[Category("Images"),
		Description("This property defines how the images will be named by FlexCel. You can always override the name using the GetImageInformation event."),
		DefaultValue(TImageNaming.Default)]
		public TImageNaming ImageNaming { get { return FImageProps.ImageNaming; } set { FImageProps.ImageNaming = value; } }


		/// <summary>
		/// Extra information to be added to the HTML file.
		/// </summary>
		[Category("Extra HTML Information"),
		Description("Extra information to be added to the HTML file.")]      
		public THtmlExtraInfo ExtraInfo {get{return FExtraInfo;}}

		/// <summary>
		/// Contains all the generated files by the component. Note that it might contain files not actually generated, if an error happened while trying to create them.
		/// </summary>
		[Browsable(false)]      
		public TGeneratedFiles GeneratedFiles {get{return FGeneratedFiles;}}

		#endregion

		#region Events
		/// <summary>
		/// Use this event to customize where to save the images when exporting to HTML.
		/// </summary>
		[Category("Images"),
		Description("Use this event to customize where to save the images when exporting to HTML.")]      
		public event ImageInformationEventHandler GetImageInformation;

		/// <summary>
		/// Replace this event when creating a custom descendant of FlexCelHtmlExport.
		/// </summary>
		/// <param name="e"></param>
		protected internal virtual void OnGetImageInformation(ImageInformationEventArgs e)
		{
			if (GetImageInformation!=null) GetImageInformation(this, e);
		}

		/// <summary>
		/// Use this event to save the images into other place. Note that this event only fires when saving HTML, not MTHML.
		/// </summary>
		[Category("Images"),
		Description("Use this event to save the images into other place.")]
		public event SaveImageEventHandler SaveImage;

		/// <summary>
		/// Replace this event when creating a custom descendant of FlexCelHtmlExport.
		/// </summary>
		/// <param name="e"></param>
		protected internal virtual void OnSaveImage(SaveImageEventArgs e)
		{
			if (SaveImage != null) SaveImage(this, e);
		}

		/// <summary>
		/// Override this property when creating your own <see cref="OnSaveImage"/> descendant. This method should return true
		/// if there is any event attached to it.
		/// </summary>
		public virtual bool HasSaveImageEvent
		{
			get { return SaveImage != null; }
		}

		/// <summary>
		/// Use this event to customize the fonts used in the exported file.
		/// </summary>
		[Category("Fonts"),
		Description("Use this event to customize the fonts used in the exported file.")]      
		public event HtmlFontEventHandler HtmlFont;

		/// <summary>
		/// Replace this event when creating a custom descendant of FlexCelHtmlExport.
		/// </summary>
		/// <param name="e"></param>
		protected virtual void OnHtmlFont(HtmlFontEventArgs e)
		{
			if (HtmlFont!=null) HtmlFont(this, e);
		}

        /// <summary>
        /// Use this event to customize how a named range if exported to the HTML file. Note that for this event to be called,
        /// you first need to set <see cref="ExportNamedRanges"/> = true. If you want to change the id that will be exported or
        /// exclude certain named from being exported, you can do so here.
        /// </summary>
        [Category("Named Ranges"),
        Description("Use this event to customize how a named range if exported to the HTML file.")]
        public event NamedRangeExportEventHandler NamedRangeExport;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelHtmlExport.
        /// </summary>
        /// <param name="e"></param>
        protected internal virtual void OnNamedRangeExport(NamedRangeExportEventArgs e)
        {
            if (NamedRangeExport != null) NamedRangeExport(this, e);
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

		#region Utility

		internal static string ReplaceMacros(string s, Dictionary<string, string> Macros)
		{
			StringBuilder Result = new StringBuilder(s.Length);
			int i = 0; 
			while (i < s.Length) 
			{
				if (i < s.Length + 3 &&  s[i] == '<' && s[i+1] == '#')
				{
					int k = i+2;
					while (k < s.Length && s[k] != '>') k++;
					if ( k < s.Length)
					{
						string tag = s.Substring(i+2, k - (i+2));
						if (Macros.ContainsKey(tag))
						{
							Result.Append(Macros[tag]);
							i= k +1;
							continue;
						}
					}
				}
				else
				{
					Result.Append(s[i]);
				}

				i++;
			}

			return Result.ToString();
		}

		internal static Uri GetMimeLocation(string s)
		{
			return THtmlEngine.GetFileUrl(Path.GetFileName(s));
		}

		internal static Uri GetMimeMainLocation(string s)
		{
			return GetMimeLocation("Main_" + Path.GetFileNameWithoutExtension(s));
		}

		#endregion

		#region Standard Export Methods
		/// <summary>
		/// Exports the active sheet of the associated xls workbook to a TextWriter. Note that you need to supply the stream for saving the images 
		/// in the <see cref="GetImageInformation"/> event (if you want to save the images).
		/// </summary>
		/// <param name="html">TextWriter where the result will be written.</param>
		/// <param name="fileName">FileName used to generate the supporting files. If you leave it null, no images will be saved as there will not be filename for them.</param>
		/// <param name="css">Use this parameter to store all CSS information in an external file. Set it to null
		/// if you want to store the CSS inside the HTML file. If you want to share the CSS between multiple files, make sure
		/// you pass the same css parameter to all Export calls.</param>
		public void Export(TextWriter html, string fileName, TCssInformation css)
		{
			GeneratedFiles.Clear();
			Workbook.Recalc(false);
			ExportToHtml(html, css, null, fileName, null, null, true);
		}	

		/// <summary>
		/// Exports the active sheet of the the associated xls workbook to a file. CSS will be saved internally in the file.
		/// </summary>
		/// <param name="htmlFileName">File name of the html file to be created.</param>
		/// <param name="relativeImagePath">Folder where images will be stored, relative to the main file. 
		/// If for example htmlFileName is "c:\reports\html\index.htm" and relativeImagePath is "images", images will be saved
		/// in folder "c:\reports\html\images". If this parameter is null or empty, images will be saved in the same folder as the html file.</param>
		public void Export(string htmlFileName, string relativeImagePath)
		{
			GeneratedFiles.Clear();
			Workbook.Recalc(false);
			DoExport(htmlFileName, relativeImagePath, null, null);
		}

		/// <summary>
		/// Exports the active sheet of the the associated xls workbook to a file.
		/// </summary>
		/// <param name="htmlFileName">File name of the html file to be created.</param>
		/// <param name="relativeImagePath">Folder where images will be stored, relative to the main file. 
		/// If for example htmlFileName is "c:\reports\html\index.htm" and relativeImagePath is "images", images will be saved
		/// in folder "c:\reports\html\images". If this parameter is null or empty, images will be saved in the same folder as the html file.</param>
		/// <param name="css">Use this parameter to store all CSS information in an external file. Set it to null
		/// if you want to store the CSS inside the HTML file. If you want to share the CSS between multiple files, make sure
		/// you pass the same css parameter to all Export calls.</param>
		public void Export(string htmlFileName, string relativeImagePath, TCssInformation css)
		{
			GeneratedFiles.Clear();
			Workbook.Recalc(false);
			DoExport(htmlFileName, relativeImagePath, css, null);
		}

		/// <summary>
		/// Exports the active sheet of the the associated xls workbook to a file.
		/// </summary>
		/// <param name="htmlFileName">File name of the html file to be created.</param>
		/// <param name="relativeImagePath">Folder where images will be stored, relative to the main file. 
		/// If for example htmlFileName is "c:\reports\html\index.htm" and relativeImagePath is "images", images will be saved
		/// in folder "c:\reports\html\images". If this parameter is null or empty, images will be saved in the same folder as the html file.</param>
		/// <param name="relativeCssFileName">Name for the Css file, with a path relative to the htmlPath. 
		/// If you set it to null, no css file created and the css will be stored inside each HTML file.
		/// If for example htmlpath is "c:\reports" and relativeCssFileName is "css\data.css" the
		/// css file will be saved in "c:\reports\css\data.css"</param>
		public void Export(string htmlFileName, string relativeImagePath, string relativeCssFileName)
		{
			GeneratedFiles.Clear();
			Workbook.Recalc(false);
			TCssInformation css = null;
			TextWriter fcss = null;
			string cssFileName = null;
			try
			{
				try
				{
					if (relativeCssFileName != null)
					{
						string htmlPath = Path.GetDirectoryName(htmlFileName);
						cssFileName = Path.Combine(htmlPath, relativeCssFileName);

						Directory.CreateDirectory(Path.GetDirectoryName(cssFileName));

						fcss = CreateStreamWriter(cssFileName, THtmlFileType.Css);
						css = new TCssInformation(fcss, relativeCssFileName);
					}

					DoExport(htmlFileName, relativeImagePath, css, null);
				}
				finally
				{
					if (fcss != null) fcss.Close();
				}
			}
            catch (IOException)
            {
                //Don't delete the file in an io exception. It might be because allowoverwritefiles was false, and the file existed.
                throw;
            }
			catch 
			{
                FlxUtils.TryDelete(cssFileName);
				throw;
			}
		}

		private void DoExport(string htmlFileName, string relativeImagePath, TCssInformation css, TSheetSelector SheetSelector)
		{
            bool HtmlCreated = false;
			try
			{
				using (TextWriter fhtm = CreateStreamWriter(htmlFileName, THtmlFileType.Html))
				{
                    HtmlCreated = true;
					if (fhtm == null) return;
					ExportToHtml(fhtm, css, null, htmlFileName, relativeImagePath, SheetSelector, true);
				}
			}
            catch (IOException)
            {
                //Don't delete the file in an io exception. It might be because allowoverwritefiles was false, and the file existed.
                if (HtmlCreated) FlxUtils.TryDelete(htmlFileName);
                throw;
            }
            catch 
			{
                FlxUtils.TryDelete(htmlFileName);
                throw;
			}
		}

		/// <summary>
		/// This method will export all the visible sheets on an xls file to an html file, writing each sheet in a different file.
		/// This is equivalent to calling Export on every sheet.
		/// </summary>
		/// <param name="htmlPath">Path where html files will be stored. (one per sheet in the workbook)</param>
		/// <param name="htmlFileNamePrefix">This is a string that will be added to every file generated, at the beggining of the filename.
		/// For example, if htlmFileNamePrefix = "test_" and htmlFileNamePostfix = ".html", "sheet1" will be exported as "test_sheet1.html".
		/// If generating a single file (MHTML format), htmlFileNamePrefix + htmlFileNamePostFix will be used, without a sheet name.</param>
		/// <param name="htmlFileNamePostfix">This is a string that will be added to every file generated, at the end of the filename. Make sure you include the extension here.
		/// For example, if htlmFileNamePrefix = "test_" and htmlFileNamePostfix = ".html", "sheet1" will be exported as "test_sheet1.html"
		/// If generating a single file (MHTML format), htmlFileNamePrefix + htmlFileNamePostFix will be used, without a sheet name.</param>
		/// <param name="relativeImagePath">Folder where images will be stored, relative to the htmlPath. 
		/// If for example htmlPath is "c:\reports\html" and relativeImagePath is "images", images will be saved
		/// in folder "c:\reports\html\images". If this parameter is null or empty, images will be saved in the same folder as the html files.</param>
		/// <param name="relativeCssFileName">Name for the Css file, with a path relative to the htmlPath. Note that the css will be shared among all sheets, so only one file will be created.
		/// if you set it to null, no css file created and the css will be stored inside each HTML file. It is recommended that you provide an extenal name here,
		/// so the CSS is shared webpages are smaller. If for example htmlpath is "c:\reports" and relativeCssFileName is "css\data.css" the
		/// css file will be saved in "c:\reports\css\data.css"</param>
		/// <param name="sheetSelector">Information about how to draw the tabs that will allow you to switch between the sheets.
		/// Set it to null if you do not want to include a sheet selector.</param>
		public void ExportAllVisibleSheetsAsTabs(string htmlPath, string htmlFileNamePrefix,string htmlFileNamePostfix, 
			string relativeImagePath, string relativeCssFileName, TSheetSelector sheetSelector)
		{
			GeneratedFiles.Clear();
			Workbook.Recalc(false);
			Canceled = false;
			TCssInformation css =  null;
			TextWriter fcss = null;
			string cssFileName = null;
			try
			{
				try
				{
					if (relativeCssFileName != null)
					{
						cssFileName = Path.Combine(htmlPath, relativeCssFileName);

						Directory.CreateDirectory(Path.GetDirectoryName(cssFileName));

						fcss = CreateStreamWriter(cssFileName, THtmlFileType.Css);
						css = new TCssInformation(fcss, relativeCssFileName);
					}

					int SaveActiveSheet = Workbook.ActiveSheet;
					try
					{
						try
						{
							if (sheetSelector != null)
							{
								for (int sheet = 1; sheet <= Workbook.SheetCount; sheet++)
								{
									if (Canceled) return;
									Workbook.ActiveSheet = sheet;
									if (Workbook.SheetVisible != TXlsSheetVisible.Visible) continue;

									sheetSelector.AddLink(htmlFileNamePrefix + THtmlEngine.EncodeFileName(Workbook.SheetName) + htmlFileNamePostfix);
								}				
							}


							TextWriter html = GetStreamWriterForTabs(Path.Combine(htmlPath, htmlFileNamePrefix + htmlFileNamePostfix), THtmlFileType.Html);
							try
							{
								TMimeWriter DedicatedMime =  null;
								if (FHtmlFileFormat == THtmlFileFormat.MHtml) 
								{
									if (html == null) return; //sharing violation.
									DedicatedMime = new TMimeWriter();
									DedicatedMime.CreateMultiPartMessage(html, TMultipartType.Related, "text/html", GetMimeLocation(htmlFileNamePrefix + htmlFileNamePostfix));
								}

								if (Canceled) return;
								for (int sheet = 1; sheet <= Workbook.SheetCount; sheet++)
								{
									Workbook.ActiveSheet = sheet;
									if (Canceled) return;
									if (Workbook.SheetVisible != TXlsSheetVisible.Visible) continue;							
								
									string FileName = Path.Combine(htmlPath, htmlFileNamePrefix + THtmlEngine.EncodeFileName(Workbook.SheetName) + htmlFileNamePostfix);

									if (HtmlFileFormat == THtmlFileFormat.Html)
									{
										html = CreateStreamWriter(FileName, THtmlFileType.Html);
										if (html == null) continue; //sharing violation.
									}
									try
									{
										ExportOneSheetAsTab(html, DedicatedMime, relativeImagePath, css, FileName, sheetSelector);
									}
									finally
									{
										if (HtmlFileFormat == THtmlFileFormat.Html)
										{
                                            if (html != null)
                                            {
                                                html.Close();
                                                html = null;
                                            }
										}
									}
								}

								if (DedicatedMime != null) DedicatedMime.EndMultiPartMessage(html);
							}
							finally
							{
                                if (html != null)
                                {
                                    html.Close();
                                    html = null;
                                }
							}
						}
						finally
						{
							if (sheetSelector != null) sheetSelector.ClearLinks();
						}
					}
					finally
					{
						Workbook.ActiveSheet = SaveActiveSheet;
					}
				}
				finally
				{
                    if (fcss != null)
                    {
                        fcss.Close();
                        fcss = null;
                    }
				}
			}
            catch (IOException)
            {
                //Don't delete the file in an io exception. It might be because allowoverwritefiles was false, and the file existed.
                throw;
            }
            catch 
			{
                FlxUtils.TryDelete(cssFileName);
                throw;
			}
		}

		private TextWriter GetStreamWriterForTabs(string htmlFileName, THtmlFileType htmlFileType)
		{
			if (FHtmlFileFormat != THtmlFileFormat.MHtml) return null;

			return CreateStreamWriter(htmlFileName, htmlFileType);
		}

		private void ExportOneSheetAsTab(TextWriter html, TMimeWriter DedicatedMime, string ImagePath, TCssInformation css, string FileName, TSheetSelector SheetSelector)
		{
			ExportToHtml(html, css, DedicatedMime, FileName, ImagePath, SheetSelector, false);
		}

		/// <summary>
		/// Exports all visible sheets in an xls file one after the other in the same html file.
		/// </summary>
		/// <param name="htmlFileName">Path where the generated html file will be stored.</param>
		/// <param name="relativeImagePath">Folder where images will be stored, relative to the path in htmlFileName. 
		/// If for example htmlFileName is "c:\reports\html\test.htm" and relativeImagePath is "images", images will be saved
		/// in folder "c:\reports\html\images". If this parameter is null or empty, images will be saved in the same folder as the html file.</param>
		/// <param name="relativeCssFileName">Name for the Css file, with a path relative to the htmlFileName Path.
		/// if you set it to null, no css file created and the css will be stored inside each HTML file. 
		/// If for example htmlFileName is "c:\reports\test.htm" and relativeCssFileName is "css\data.css" the
		/// css file will be saved in "c:\reports\css\data.css"</param>
		/// <param name="sheetSeparator">An HTML string to write between all the different sheets being exported. You can use for example &lt;hr /&gt; here
		/// to add an horizontal line. You can also use the special macros &lt;#SheeName&gt;, &lt;#SheeCount&gt; and &lt;#SheePos&gt; here to enter the current sheet name, 
		/// the number of sheets and the current sheet respectively. The macros are case insensitive, you can enter them in any combination of upper and lower case.</param>
		public void ExportAllVisibleSheetsAsOneHtmlFile(string htmlFileName,
			string relativeImagePath, string relativeCssFileName, string sheetSeparator)
		{
			GeneratedFiles.Clear();
			Workbook.Recalc(false);
            bool HtmlCreated = false;
			try
			{
				using (TextWriter fhtm = CreateStreamWriter(htmlFileName, THtmlFileType.Html))
				{
                    HtmlCreated = true;
					if (fhtm == null) return;
					DoExportAllVisibleSheetsAsOneHtmlFile(fhtm, htmlFileName, relativeImagePath, relativeCssFileName, sheetSeparator);
				}
			}
            catch (IOException)
            {
                //Don't delete the file in an io exception. It might be because allowoverwritefiles was false, and the file existed.
                if (HtmlCreated) FlxUtils.TryDelete(htmlFileName);
                throw;
            }
            catch
			{
                FlxUtils.TryDelete(htmlFileName);
				throw;
			}
		}


		private string ReplaceSeparatorMacros(int i, string Separator)
		{
			Dictionary<string, string> sd = new Dictionary<string,string>(StringComparer.InvariantCultureIgnoreCase);
			sd.Add("sheetpos", i.ToString(CultureInfo.CurrentCulture));
			sd.Add("sheetname", Workbook.GetSheetName(i));
			sd.Add("sheetcount", Workbook.SheetCount.ToString(CultureInfo.CurrentCulture));

			return ReplaceMacros(Separator, sd);
		}

		/// <summary>
		/// Exports all visible sheets in an xls file one after the other in the same html stream.
		/// </summary>
		/// <param name="html">Stream where the HTML file will be saved.</param>
		/// <param name="htmlFileName">The filename itself won't be used (since the html file is saved to a stream), but it will tell the path where the generated extra files (for example images) will be stored, and also the way images are named.</param>
		/// <param name="relativeImagePath">Folder where images will be stored, relative to the path in htmlFileName. 
		/// If for example htmlFileName is "c:\reports\html\test.htm" and relativeImagePath is "images", images will be saved
		/// in folder "c:\reports\html\images". If this parameter is null or empty, images will be saved in the same folder as the html file.</param>
		/// <param name="relativeCssFileName">Name for the Css file, with a path relative to the htmlFileName Path.
		/// if you set it to null, no css file created and the css will be stored inside each HTML file. 
		/// If for example htmlFileName is "c:\reports\test.htm" and relativeCssFileName is "css\data.css" the
		/// css file will be saved in "c:\reports\css\data.css"</param>
		/// <param name="sheetSeparator">An HTML string to write between all the different sheets being exported. You can use for example &lt;hr /&gt;&lt;p&gt;Sheet &lt;#SheetName&gt;&lt;/p&gt; here
		/// to add an horizontal line. You can also use the special macros &lt;#SheetName&gt;, &lt;#SheetCount&gt; and &lt;#SheetPos&gt; here to enter the current sheet name, 
		/// the number of sheets and the current sheet respectively. The macros are case insensitive, you can enter them in any combination of upper and lower case.</param>
		public void ExportAllVisibleSheetsAsOneHtmlFile(TextWriter html, string htmlFileName,
			string relativeImagePath, string relativeCssFileName, string sheetSeparator)
		{
			GeneratedFiles.Clear();
			Workbook.Recalc(false);
			DoExportAllVisibleSheetsAsOneHtmlFile(html, htmlFileName, relativeImagePath, relativeCssFileName, sheetSeparator);
		}

		private void DoExportAllVisibleSheetsAsOneHtmlFile(TextWriter html, string htmlFileName,
			string relativeImagePath, string relativeCssFileName, string sheetSeparator)
		{
			Canceled = false;
			TextWriter fcss = null;
			string cssFileName = null;
			try
			{
				try
				{
					if (relativeCssFileName != null)
					{
						cssFileName = Path.Combine(Path.GetDirectoryName(htmlFileName), relativeCssFileName);

						Directory.CreateDirectory(Path.GetDirectoryName(cssFileName));

						fcss = CreateStreamWriter(cssFileName, THtmlFileType.Css);
					}

					TPartialExportState ExportState = new TPartialExportState(fcss, relativeCssFileName);

					int SaveActiveSheet = Workbook.ActiveSheet;
					try
					{
						for (int sheet = 1; sheet <= Workbook.SheetCount; sheet++)
						{
							if (Canceled) return;
							Workbook.ActiveSheet = sheet;
							if (Workbook.SheetVisible != TXlsSheetVisible.Visible) continue;

							PartialExportAdd(ExportState, htmlFileName, relativeImagePath);
						}
				
						if (Canceled) return;
						ExportState.SaveFullHeaders(html, htmlFileName);
						ExportState.StartBody(html);

						for (int i = 1; i <= ExportState.BodyCount; i++)
						{
							if (Canceled) return;
							if (i > 1 && sheetSeparator != null) html.WriteLine(ReplaceSeparatorMacros(i, sheetSeparator));
							ExportState.SaveBody(html, i, relativeImagePath);
						}
						ExportState.EndHtmlFile(html);

					}
					finally
					{
						Workbook.ActiveSheet = SaveActiveSheet;
					}
				}
				finally
				{
                    if (fcss != null)
                    {
                        fcss.Close();
                        fcss = null;
                    }
				}
			}
            catch (IOException)
            {
                //Don't delete the file in an io exception. It might be because allowoverwritefiles was false, and the file existed.
                throw;
            }
            catch 
			{
                FlxUtils.TryDelete(cssFileName);
				throw;
			}
		}
	

		#endregion

		#region Partial Export Methods
		/// <summary>
		/// This method is designed to export a part of the HTML inside other HTML file.
		///  You should call this method once for every file you wish to export, 
		///  in order to let FlexCel consolidate the needed CSS classes and information into the partialExportState parameter.
		///  Once you have "filled" the partialExportState object with all the information on files you wish to export, you can call the methods
		///  in partialExportState to write the CSS or body parts of the HTML file.
		/// <p>Note that this overload version will automatically save the images to disk when needed. If for any reason you do not want
		/// to save the images, use the overload with SaveImagesToDisk parameter.</p>
		/// </summary>
		/// <remarks>This method does *not* clear the generated files array, so you have more control over it. If you want to use it, you need to call
		/// the clear method yourself before the first PartialExportAdd.</remarks>
		/// <param name="partialExportState">Object where we will collect all the css and needed information on the files being exported. Call this method
		/// once for every sheet you want to export, using the same partialExportState parameter.</param>
		/// <param name="htmlFileName">Name of the stored html file. This affects the path where images will be saved, and also the way images will be named.</param>
		/// <param name="relativeImagePath">Path relative to the path of htmlFileName, where images will be stored.</param>
		public void PartialExportAdd(TPartialExportState partialExportState, string htmlFileName, string relativeImagePath)
		{
			PartialExportAdd(partialExportState, htmlFileName, relativeImagePath, true);
		}

		/// <summary>
		/// This method is designed to export a part of the HTML inside other HTML file.
		///  You should call this method once for every file you wish to export, 
		///  in order to let FlexCel consolidate the needed CSS classes and information into the partialExportState parameter.
		///  Once you have "filled" the partialExportState object with all the information on files you wish to export, you can call the methods
		///  in partialExportState to write the CSS or body parts of the HTML file.
		/// </summary>
		/// <param name="partialExportState">Object where we will collect all the css and needed information on the files being exported. Call this method
		/// once for every sheet you want to export, using the same partialExportState parameter.</param>
		/// <param name="htmlFileName">Name of the stored html file. This affects the path where images will be saved, and also the way images will be named.</param>
		/// <param name="relativeImagePath">Path relative to the path of htmlFileName, where images will be stored.</param>
		/// <param name="SaveImagesToDisk">When true, any image found in the file will be saved to disk (to relative imagepath) as the file is added.
		/// If false, the links for the image will be written but no image will be saved to disk, you have to do this yourself. Note that when 
		/// saving MHTML files this parameter doesn't matter, since images will not be saved to disk anyway.</param>
		/// <remarks>This method does *not* clear the generated files array, so you have more control over it. If you want to use it, you need to call
		/// the clear method yourself before the first PartialExportAdd.</remarks>
		public void PartialExportAdd(TPartialExportState partialExportState, string htmlFileName, string relativeImagePath, bool SaveImagesToDisk)
		{
			TXlsCellRange[] FinalPrintRange = GetFinalPrintRange();

			TExportHtmlCache Cache = new TExportHtmlCache(partialExportState.CssInfo, null, this);
			bool SaveImgs = SaveImagesToDisk && HtmlFileFormat == THtmlFileFormat.Html;

			partialExportState.AddSheet(this, FWorkbook, htmlFileName, FWorkbook.ActiveSheet, FinalPrintRange, Cache,
				new TImageInformation(FImageProps, FSavedImagesFormat), Progress, ExtraInfo, this, SaveImgs);

            THtmlEngine.SearchUsedXF(FWorkbook, FinalPrintRange, Cache);                
            
			
			partialExportState.Engine.EngineRuns++; //We need to ensure an unique imageId to each PartialExportAdd call.
			partialExportState.Engine.ExtractImages(Workbook, Cache, htmlFileName, relativeImagePath, SaveImgs);
		}

		#endregion

		#region Implementation

		private TextWriter CreateStreamWriter(string FileName, THtmlFileType htmlFileType)
		{
			GeneratedFiles.Add(FileName, htmlFileType);

			FileMode fm=FileMode.CreateNew;
			if (AllowOverwritingFiles) fm=FileMode.Create;

			FileStream fs = null;
			try
			{
				fs = new FileStream(FileName, fm);
			}
			catch (IOException ex)
			{
				fs = null;
				if (!IgnoreSharingViolations || !AllowOverwritingFiles) throw;
				if (TSaveAndHandleSharingViolation.CheckRethrow(ex)) throw;
			}
			if (fs == null) return null;
			return new StreamWriter(fs);
		}

		private void ExportToHtml(TextWriter html, TCssInformation css, TMimeWriter DedicatedMime, string FileName, string ImagePath, TSheetSelector SheetSelector, bool ResetCancelded)
		{
			if (ResetCancelded) 
			{
				Canceled = false;
			}
			TXlsCellRange[] FinalPrintRange = GetFinalPrintRange();
			ConvertToHtml(html, css, FinalPrintRange, FileName, DedicatedMime, ImagePath, SheetSelector);
		}

		private TXlsCellRange[] GetFinalPrintRange()
		{
            TXlsCellRange[] FinalPrintRange = new TXlsCellRange[] { (TXlsCellRange)FPrintRange.Clone() };

			if (FPrintRange.Top <= 0 || FPrintRange.Left <= 0 || FPrintRange.Bottom <= 0 || FPrintRange.Right <= 0)
			{
				FlexCelRender Renderer = new FlexCelRender();

				using (TGraphicCanvas GrCanvas = new TGraphicCanvas())
				{
					Renderer.SetCanvas(GrCanvas.Canvas);
					Renderer.Workbook = Workbook;
					Renderer.CreateFontCache();
					try
					{
						FinalPrintRange = Renderer.InternalCalcPrintArea(FPrintRange);
					}
					finally
					{
						Renderer.DisposeFontCache();
					}
				}
			}
			return FinalPrintRange;
		}

        private void ConvertToHtml(TextWriter html, TCssInformation css, TXlsCellRange[] FinalPrintRange, string FileName, TMimeWriter DedicatedMime, string ImagePath, TSheetSelector SheetSelector)
		{
			Progress.Clear(Workbook.ActiveSheet);
			TMimeWriter MimeWriter = null;
			if (FHtmlFileFormat == THtmlFileFormat.MHtml)
			{
				MimeWriter = DedicatedMime == null? new TMimeWriter(): DedicatedMime;
			}

			THtmlEngine engine = new THtmlEngine(FClassPrefix, FHtmlVersion, FHtmlFileFormat, MimeWriter, FHidePrintObjects, FHtmlFixes, VerticalTextAsImages,
				new TImageInformation(FImageProps, FSavedImagesFormat), HeadingWidth, HeadingStyle, UseContentId, ExportNamedRanges, this);

			engine.Init(Workbook);

			TExportHtmlCache Cache = new TExportHtmlCache(css, SheetSelector, this);
			Progress.SetTotalRows(TPartialExportState.GetRowsInRange(FinalPrintRange));

			engine.ExtractImages(Workbook, Cache, FileName, ImagePath, HtmlFileFormat == THtmlFileFormat.Html);

			if (MimeWriter != null)
			{
				if (DedicatedMime == null)
				{
					MimeWriter.CreateMultiPartMessage(html, TMultipartType.Related, "text/html", GetMimeMainLocation(FileName));
				}
				MimeWriter.AddPartHeader(html, "text/html", TContentTransferEncoding.QuotedPrintable, GetMimeLocation(FileName), null, html.Encoding.HeaderName);
			}

			THtmlEngine.SearchUsedXF(Workbook, FinalPrintRange, Cache);
			engine.WriteHeader(Workbook, html, ExtraInfo, Cache.CssInfo, Cache.Images.Count > 0, SheetSelector);
			
			engine.StartBody(html, ExtraInfo);

			try
			{
				if (SheetSelector != null)
				{
					SheetSelector.SetInternals(engine, html);

					SheetSelector.DrawSelector(Workbook, TSheetSelectorPosition.Top);
					SheetSelector.DrawSelector(Workbook, TSheetSelectorPosition.Left);
					engine.WriteLn(html, "<div class = 'sheet_content'>");
				}

                engine.WriteBody(Workbook, html, FinalPrintRange, Cache, FProgress, FileName, ImagePath, 
                    HtmlFileFormat == THtmlFileFormat.Html, ExtraInfo);
				if (SheetSelector != null) 
				{
					engine.WriteLn(html, "</div>");
					SheetSelector.DrawSelector(Workbook, TSheetSelectorPosition.Right);
					SheetSelector.DrawSelector(Workbook, TSheetSelectorPosition.Bottom);
				}
			}
			finally
			{
				if (SheetSelector != null) SheetSelector.ClearInternals();
			}

			engine.EndBody(html, ExtraInfo);
			engine.WriteEndDoc(html);

			if (MimeWriter != null)
			{
				MimeWriter.EndPart(html);
				engine.SaveMHTMLImages(Workbook, html, Cache, MimeWriter);
				
				if (DedicatedMime == null) MimeWriter.EndMultiPartMessage(html);  //If dedicatedmime is not null, it should be the one to close it.
			}
		}

		#endregion

		#region IExportEvents Members
		/// <summary>
		/// This method is for internal use.
		/// </summary>
		/// <param name="e"></param>
		public void DoHtmlFont(HtmlFontEventArgs e)
		{
			OnHtmlFont(e);
		}

		#endregion
	}

	#endregion

	#region PartialExportState

	/// <summary>
	/// This class is used to save the needed information to partially export a file.
	/// </summary>
	public class TPartialExportState
	{
		private List<TSheetState> Sheets;
		internal TCssInformation CssInfo;
		internal TMimeWriter MimeWriter;
		internal THtmlEngine Engine;
		internal THtmlExtraInfo ExtraInfo;

        FlexCelHtmlExportProgress Progress;

		/// <summary>
		/// Creates a new instance of TPartialExportState.
		/// </summary>
		/// <param name="aCssData">TextWriter where an external CSS file will be stored. If null, no CSS file will be created, even when a link to a external file might.</param>
		/// <param name="aCssUrl">URL of the css file that will be linked to this file. If null, all css information will be stored inside the html file.</param>
		public TPartialExportState(TextWriter aCssData, string aCssUrl)
		{
            CssInfo = new TCssInformation(aCssData, aCssUrl);
            Sheets = new List<TSheetState>();
        }

        internal void AddSheet(FlexCelHtmlExport htmlExport, ExcelFile aXls, string aHtmlFileName, int aActiveSheet, 
			TXlsCellRange[] aCellRange, TExportHtmlCache aCache, TImageInformation aImageInfo, FlexCelHtmlExportProgress aProgress,
			THtmlExtraInfo aExtraInfo, FlexCelHtmlExport aParent, bool aSaveImagesToDisk)
		{
			if (Engine == null)
			{
				MimeWriter = htmlExport.HtmlFileFormat == THtmlFileFormat.MHtml ? new TMimeWriter() : null;
				Engine = new THtmlEngine(htmlExport.ClassPrefix, htmlExport.HtmlVersion, htmlExport.HtmlFileFormat, MimeWriter, 
					htmlExport.HidePrintObjects, htmlExport.FHtmlFixes, htmlExport.VerticalTextAsImages, aImageInfo, 
					htmlExport.HeadingWidth, htmlExport.HeadingStyle, htmlExport.UseContentId, htmlExport.ExportNamedRanges, aParent);
                Progress = aProgress;
				Progress.Clear(aXls.ActiveSheet);
				ExtraInfo = aExtraInfo;
			}
			Sheets.Add(new TSheetState(aXls, aHtmlFileName, aActiveSheet, aCellRange, aCache, aSaveImagesToDisk));
		}

		private TSheetState Sheet(int index)
		{
			return Sheets[index];
		}

		
		/// <summary>
		/// Use this method to output the CSS information on this object to the header of an HTML page. If you are using en external
		/// StyleSheet, this method will output a link to it, or if you are using an internal one it will output the actual classes.
		/// </summary>
		/// <param name="writer">Writer where you are going to write the information.</param>
		public void SaveCss(TextWriter writer)
		{
			Engine.WriteCss(writer, CssInfo, null);
		}

		/// <summary>
		/// This method is a middle ground between <see cref="SaveCss"/> and <see cref="SaveFullHeaders"/>.
		/// It will output only the headers that you need to add to an existing HTML file in order to include the body inthe body part.
		/// This means that the tags like &lt;html&gt; are not included.
		/// </summary>
		/// <param name="writer"></param>
		public void SaveRelevantHeaders(TextWriter writer)
		{
			Engine.WriteCss(writer, CssInfo, null);
			bool HasImages = false;
			foreach (TSheetState sh in Sheets)
			{
				if (sh.Cache.Images.Count > 0) {HasImages = true; break;}
			}
			if (Engine.HtmlFixes.IE6TransparentPngSupport) Engine.AddIe6TransparentPngFix(writer, HasImages);
		}

		/// <summary>
		/// This method will output the full HTML headers needed to create an HTML file with the information in this object.
		/// If you wish to mix the output of the file with existing headers, you can use <see cref="SaveRelevantHeaders"/> instead to get only
		/// the relevant information to mix in the headers, or <see cref="SaveCss"/> to get only the CSS classes that need to be put in the header.
		/// </summary>
		/// <param name="writer">Writer where you are going to write the information.</param>
		/// <param name="htmlFileName">File name of the file you are generating. There is no need to supply this parameter, it is only to add extra
		/// information to the generated file. You can leave it null.</param>
        public void SaveFullHeaders(TextWriter writer, string htmlFileName)
        {
            if (MimeWriter != null)
            {
                if (htmlFileName == null || htmlFileName.Trim().Length == 0) htmlFileName = "report.mht";
                MimeWriter.CreateMultiPartMessage(writer, TMultipartType.Related, "text/html", FlexCelHtmlExport.GetMimeMainLocation(htmlFileName));
                MimeWriter.AddPartHeader(writer, "text/html", TContentTransferEncoding.QuotedPrintable, FlexCelHtmlExport.GetMimeLocation(htmlFileName), null, writer.Encoding.HeaderName);
            }

            bool HasImages = false;
            foreach (TSheetState sh in Sheets)
            {
                if (sh.Cache.Images.Count > 0) { HasImages = true; break; }
            }
            Engine.WriteHeader(null, writer, ExtraInfo, CssInfo, HasImages, null);
        }

		/// <summary>
		/// Number of parts added to this object.
		/// </summary>
		public int BodyCount
		{
			get
			{
				return Sheets.Count;
			}
		}

        /// <summary>
        /// Starts writing a body declaration. After calling this method, you should call <see cref="SaveBody"/> for the
        /// parts you want to save, and end up with a call to <see cref="EndHtmlFile"/>
        /// </summary>
        /// <param name="writer">TextWriter where we are going to save the results.</param>
		public void StartBody(TextWriter writer)
		{
            if (Progress == null) return;
			Engine.StartBody(writer, ExtraInfo);
		}

        /// <summary>
        /// Writes the "&lt;/body&gt;" end tag in the html file and the head/html end tags. It also finalizes the parts when saving to MHTML.
        /// </summary>
        /// <param name="writer">TextWriter where we are going to save the results.</param>
		public void EndHtmlFile(TextWriter writer)
		{
			if (Engine == null) return;
			Engine.EndBody(writer, ExtraInfo);
			Engine.WriteEndDoc(writer);
			
			if (MimeWriter != null)
			{
				MimeWriter.EndPart(writer);
				foreach (TSheetState sh in Sheets)
				{
					int SaveActiveSheet = sh.Xls.ActiveSheet;
					try
					{
						sh.Xls.ActiveSheet = sh.ActiveSheet;
						Engine.SaveMHTMLImages(sh.Xls, writer, sh.Cache, MimeWriter);
					}
					finally
					{
						sh.Xls.ActiveSheet = SaveActiveSheet;
					}
				}
				MimeWriter.EndMultiPartMessage(writer);
			}
		}
		

		/// <summary>
		/// Use this method to output the body information on this object to an HTML page.
		/// </summary>
		/// <param name="writer">Writer where you are going to write the information.</param>
		/// <param name="index">Index of the part that you wish to write. It must be 1 &lt;= index &lt;= <see cref="BodyCount"/></param>
        /// <param name="relativeImagePath">Image path relative to the main file where the images will be saved. Note that this path 
        /// <b>does not apply to normal images.</b> This is used for example to save the rotated text as images if this option is enabled.</param>
		public void SaveBody(TextWriter writer, int index, string relativeImagePath)
		{
            if (Engine == null) return;
			
			TSheetState sh = Sheet(index - 1);
			Progress.Clear(sh.ActiveSheet);
			Progress.SetTotalRows(GetRowsInRange(sh.CellRange));

			int SaveActiveSheet = sh.Xls.ActiveSheet;
			try
			{
				sh.Xls.ActiveSheet = sh.ActiveSheet;
				Engine.Init(sh.Xls);

				Engine.WriteBody(sh.Xls, writer, sh.CellRange, sh.Cache, Progress, sh.HtmlFileName, 
                    relativeImagePath, sh.SaveImagesToDisk, ExtraInfo);
			}
			finally
			{
				sh.Xls.ActiveSheet = SaveActiveSheet;
			}
		}

        internal static int GetRowsInRange(TXlsCellRange[] ranges)
        {
            int Result = 0;
            foreach (TXlsCellRange rng in ranges)
            {
                Result += rng.RowCount;
            }
            return Result;
        }

        /// <summary>
        /// Returns one of the images for one of the saved sheets.
        /// </summary>
        /// <param name="bodyIndex">Index of the document where you want to retrieve the images. (1 based)</param>
        /// <param name="imageIndex">Image index in the document. (1 based)</param>
        /// <returns></returns>
        public Image GetImage(int bodyIndex, int imageIndex)
        {
			TSheetState sh = Sheet(bodyIndex - 1);

            TShapeProperties ShProp = sh.Xls.GetObjectProperties(imageIndex, true);
            if (!ShProp.Print || !ShProp.Visible || ShProp.ObjectType == TObjectType.Comment) return null;

            RectangleF Dimensions;
            PointF Origin;
            Size ImageSizePixels;

			TImageProps ImageProps = Engine.Images.Props;
            return sh.Xls.RenderObject(imageIndex, ImageProps.ImageResolution, ShProp, ImageProps.SmoothingMode, ImageProps.InterpolationMode,
                ImageProps.AntiAliased, true, ImageProps.ImageBackground, out Origin, out Dimensions, out ImageSizePixels);
        }

	}

	internal class TSheetState
	{
		internal ExcelFile Xls;
		internal int ActiveSheet;
		internal TXlsCellRange[] CellRange;
        internal TExportHtmlCache Cache;
        internal string HtmlFileName;
        internal bool SaveImagesToDisk;

		internal TSheetState(ExcelFile aXls, string aHtmlFileName, int aActiveSheet,  TXlsCellRange[] aCellRange, TExportHtmlCache aCache, bool aSaveImagesToDisk)
		{
			Xls = aXls;
            HtmlFileName = aHtmlFileName;
			ActiveSheet = aActiveSheet;
			CellRange = aCellRange;
			Cache = aCache;
            SaveImagesToDisk = aSaveImagesToDisk;
		}
	}

	#endregion

	#region SheetSelector

	internal class TEngineState
	{
		internal THtmlEngine Engine;
		internal TextWriter html;

		internal TEngineState(THtmlEngine aEngine, TextWriter aHtml)
		{
			Engine = aEngine;
			html = aHtml;
		}
	}

	/// <summary>
	/// Abstract class to implement a Sheet Selector. Derive from this class for example to implement
	/// tabs with images. For a standard implementation using CSS Tabs and divs, use <see cref="TStandardSheetSelector"/>
	/// </summary>
	public abstract class TSheetSelector
	{
		#region Privates
		private THtmlEngine Engine;
		private TextWriter html;
		private Stack<TEngineState> EngineState;
        private readonly List<string> FLinks;
		private readonly TSheetSelectorPosition FSheetSelectorPosition;

        /// <summary>
        /// An enumerator defining all the positions where the SheetSelector will be drawn.Read it to know where to draw the selector.
        /// </summary>
		protected TSheetSelectorPosition SheetSelectorPosition {get{return FSheetSelectorPosition;}}

        /// <summary>
        /// A list of links that should go in the sheet selector, one per tab. Use them when creating your
        /// own sheet selector to know where to point the link in the tabs to.
        /// </summary>
		protected IList<string> Links{get {return FLinks;}}
		#endregion

		#region Constructor
		/// <summary>
		/// Contructs a new TSheetSelector instance.
		/// </summary>
		/// <param name="position">Position where the Sheet selector will be placed. Remember you can 'or' together many values
		/// on the enumeration to for example have a selector at the top and at the bottom.</param>
		protected TSheetSelector(TSheetSelectorPosition position)
		{
			FSheetSelectorPosition = position;
			FLinks = new List<string>();
			EngineState = new Stack<TEngineState>(2);
		}
		#endregion

		#region Internals
		internal void SetInternals(THtmlEngine aEngine, TextWriter aHtml)
		{
			Engine = aEngine;
			html = aHtml;
		}

		internal void ClearInternals()
		{
			Engine = null;
			html = null;
		}

		internal void AddInternals(THtmlEngine aEngine, TextWriter aHtml)
		{
			EngineState.Push(new TEngineState(Engine, html));

			Engine = aEngine;
			html = aHtml;
		}

		internal void RestoreInternals()
		{
			TEngineState es = EngineState.Pop();
			Engine = es.Engine;
			html = es.html;
		}


		internal void AddLink(string s)
		{
			Links.Add(s);
		}

		internal void ClearLinks()
		{
			Links.Clear();
		}

		internal void DrawSelector(ExcelFile xls, TSheetSelectorPosition Reference)
		{
				int SaveActiveSheet = xls.ActiveSheet;
				try
				{
                    BeforeDrawOneSheetSelector(Reference);
                    if ((SheetSelectorPosition & Reference) != 0)
                    {
                        DrawOneSheetSelector(xls, Reference);
                    } 
                    AfterDrawOneSheetSelector(Reference);

				}
				finally
				{
					xls.ActiveSheet = SaveActiveSheet;
				}
		}

		#endregion

		#region Write Sheet Selector

		/// <summary>
		/// Use this method to write a line inside the stream when overriding this class.
		/// </summary>
		/// <param name="s"></param>
		public void WriteLn(string s)
		{
			Engine.WriteLn(html, s);
		}

		/// <summary>
		/// This method will encode an string so it is valid html. For example, it will replace "&amp;" by "&amp;amp;" in the text.
		/// </summary>
		/// <param name="s">String that you want to encode.</param>
		/// <returns>Encoded string.</returns>
		public string EncodeAsHtml(string s)
		{
			return THtmlEntities.EncodeAsHtml(s, Engine.HtmlVersion, html.Encoding);
		}

        /// <summary>
        /// Use this method to customize actions to do before the SheetSelector is drawn. 
        /// In the <see cref="TStandardSheetSelector"/> implementation, this method is used to add a table for layout
        /// if <see cref="TStandardSheetSelector.LayoutTable"/> is true.
        /// Note that this method is called once for each of the possible positions of Reference, even if you do not need to draw a selector in that position.
        /// The order in which this method will be called is: Top, Left, Right, Bottom.
        /// </summary>
        /// <param name="Reference">The position of the SheetSelector that is being created. 
        /// you can use <see cref="SheetSelectorPosition"/> to know if this is one of the selectors you need to render.</param>
        public abstract void AfterDrawOneSheetSelector(TSheetSelectorPosition Reference);

        /// <summary>
        /// Use this method to customize actions to do after the SheetSelector is drawn. 
        /// In the <see cref="TStandardSheetSelector"/> implementation, this method is used to add a table for layout
        /// if <see cref="TStandardSheetSelector.LayoutTable"/> is true.
        /// Note that this method is called once for each of the possible positions of Reference, even if you do not need to draw a selector in that position.
        /// The order in which this method will be called is: Top, Left, Right, Bottom.
        /// </summary>
        /// <param name="Reference">The position of the SheetSelector that is being created. 
        /// you can use <see cref="SheetSelectorPosition"/> to know if this is one of the selectors you need to render.</param>
        public abstract void BeforeDrawOneSheetSelector(TSheetSelectorPosition Reference);

		/// <summary>
		/// Override this method on a child class if you want to completely customize how the Sheet Selector is drawn.
		/// Normally when deriving from <see cref="TStandardSheetSelector"/> you can just change the CSS properties of this class to customize the SheetSelector, but you can use this if you want to 
		/// provide a completely different selector. You can use the <see cref="Links"/> collection to know which hyperlinks to place in each place.
        /// Note that different from <see cref="BeforeDrawOneSheetSelector"/>, this method is <b>only</b> called for the positions where the selector has to be rendered.
		/// </summary>
		/// <param name="xls">ExcelFile we are exporting. It is positioned on the sheet we are exporting, but its active sheet can be changed on this method
		/// and there is no need to restore it. It will be restored by the framework.</param>
		/// <param name="Position">Position where the sheet selector will be placed. Note that this method will always
		/// be called with only one value, but it might be called more than once if constants in the TSheetSelectorPosition enumeration are
		/// or'ed together.</param>
		public abstract void DrawOneSheetSelector(ExcelFile xls, TSheetSelectorPosition Position);

		/// <summary>
		/// This method is in charge of writing the style definitions in the header of the html file.
		/// Note that when deriving from <see cref="TStandardSheetSelector"/> you normally do not need to override this method, you can just change the CSS properties of this class.
		/// You can override this method if you want full control on how to export the classes.
		/// </summary>
		public abstract void WriteCssClasses();

		#endregion
	}

	/// <summary>
	/// Holds the styles for one of the positions of a <see cref="TStandardSheetSelector"/>.
	/// </summary>
	public class TStandardSheetSelectorStyles
	{
		#region Privates
		private string FMain;
		
		private string FActiveTab;
		private string FUnselectedTab;
		private string FUnselectedTabHover;

		private string FLinks;
		private string FActiveText;
		private string FList;

		#endregion

		#region Properties
		
		/// <summary>
		/// Style to be applied to the whole Selector.
		/// </summary>
		public string Main {get {return FMain;} set{FMain = value;}}

		/// <summary>
		/// Style to be applied to the Active tab.
		/// </summary>
		public string ActiveTab {get {return FActiveTab;} set{FActiveTab = value;}}

		/// <summary>
		/// Style to be applied to a tab when it is not the Active one.
		/// </summary>
		public string UnselectedTab {get {return FUnselectedTab;} set{FUnselectedTab = value;}}

		/// <summary>
		/// Style to be applied to an unselected tab when you hover the mouse over it.
		/// </summary>
		public string UnselectedTabHover {get {return FUnselectedTabHover;} set{FUnselectedTabHover = value;}}

		/// <summary>
		/// Style to be applied to the links in the unselected tabs. Note that the active tab does not have a link.
		/// </summary>
		public string Links {get {return FLinks;} set{FLinks = value;}}

		/// <summary>
		/// Style to be applied to the text of the selected tab.
		/// </summary>
		public string ActiveText {get {return FActiveText;} set{FActiveText = value;}}

		/// <summary>
		/// Style of the list used for the different entries in the selector. this defaults to
		/// "list-style: none", but you can use something like "list-style: upper-roman inside" to for example place roman number on each text.
		/// </summary>
		public string List {get {return FList;} set{FList = value;}}

		#endregion
	}

	/// <summary>
	/// Implements a standard sheet selector (with CSS tabs) that will allow you to change the page when exporting multiple sheets.
	/// You can customize its default behavior by altering the CSS properties, or by inheriting from it and replacing the virtual methods.
	/// If you want to create a completely new type of sheet selector, derive it from <see cref="TSheetSelector"/> instead of this class.
	/// </summary>
	public class TStandardSheetSelector: TSheetSelector
	{
		#region Privates
		private readonly TStandardSheetSelectorStyles FCssGeneral;
		private readonly TStandardSheetSelectorStyles FCssWhenTop;
		private readonly TStandardSheetSelectorStyles FCssWhenLeft;
		private readonly TStandardSheetSelectorStyles FCssWhenBottom;
		private readonly TStandardSheetSelectorStyles FCssWhenRight;

		private string FCssStyleSheetContent;
		private string FCssStyleLayoutTable;
		private readonly Dictionary<string, string> FCssTags;
        private bool FLayoutTable;
		private bool FUseSheetTabColors;
		#endregion

		#region Constructor
		/// <summary>
		/// Contructs a new TStandardSheetSelector instance.
		/// </summary>
		/// <param name="position">Position where the Sheet selector will be placed. Remember you can 'or' together many values
		/// on the enumeration to for example have a selector at the top and at the bottom.</param>
		public TStandardSheetSelector(TSheetSelectorPosition position): base(position)
		{
			FUseSheetTabColors = true;

			FCssGeneral    = new TStandardSheetSelectorStyles();
			FCssWhenLeft   = new TStandardSheetSelectorStyles();
			FCssWhenTop    = new TStandardSheetSelectorStyles();
			FCssWhenRight  = new TStandardSheetSelectorStyles();
			FCssWhenBottom = new TStandardSheetSelectorStyles();

			//CssGeneral.Main = "";
			CssWhenLeft.Main = "float: left; width: <#width>;overflow:hidden;margin-right:20pt; margin-left: 10pt;";
			CssWhenRight.Main = "float: left; width: <#width>;overflow:hidden;margin-left:20pt; margin-left: 10pt;";
			CssWhenTop.Main = "clear: both;padding:5pt 1pt 5pt 1pt;";
			CssWhenBottom.Main = "clear: both; padding: 5pt 0;";

			CssGeneral.ActiveTab     = "border: 1px solid <#bordercolor>; background: <#pagecolor>; padding: 0 0;";
			CssWhenLeft.ActiveTab   = "border:0; border-bottom: 1px solid <#bordercolor>;";
			CssWhenTop.ActiveTab    = "border-bottom-color: <#pagecolor>;display:inline-block; margin-right:1pt;float:left;height:20px;padding-top: 1px;position:relative;margin-bottom:-1px";
			CssWhenRight.ActiveTab  = "border:0; border-bottom: 1px solid <#bordercolor>";
			CssWhenBottom.ActiveTab = "border-top-color: <#pagecolor>;display:inline-block; margin-right:1pt;float:left;height:20px;padding-bottom: 1px;margin-top:-1px;position:relative;";
			
			CssGeneral.UnselectedTab    = "border: 1px solid <#bordercolor>; background: <#unselectedtabbg>; color: <#unselectedtabfg>;";
			CssWhenTop.UnselectedTab    = "display:inline-block;float:left;margin-right:1pt;height:20px;padding-top:1px;border-bottom:none;position:relative;margin-bottom:-1px";
			CssWhenBottom.UnselectedTab = "display:inline-block;float:left;margin-right:1pt;height:20px;padding-bottom:1px;margin-top:-1px;position:relative;";
			CssWhenLeft.UnselectedTab   = "border:0; border-bottom: 1px solid <#bordercolor>";
			CssWhenRight.UnselectedTab  = "border:0; border-bottom: 1px solid <#bordercolor>";
			
			CssGeneral.UnselectedTabHover = "background: <#hoverbg>; color: <#hoverfg>;";


			CssGeneral.Links = "text-decoration: none; color: <#unselectedtabfg>;padding: 0 5pt; display:block;"; //use display-block for the hyperlink using all box area.
			
			CssGeneral.List    = "list-style:none;padding: 0 0 0 10pt;margin:2px;";
			CssWhenLeft.List   = "margin-top: 50px; border: 3px solid <#bordercolor>; padding:0;";
			CssWhenTop.List    = "border-bottom: 1px solid <#bordercolor>;white-space:nowrap;height:22px;";
			CssWhenRight.List  = "margin-top: 50px;border: 3px solid <#bordercolor>; padding:0";
			CssWhenBottom.List = "border-top: 1px solid <#bordercolor>;white-space:nowrap;height:0px;";

			CssGeneral.ActiveText = "color: <#activetabfg>;padding: 0 5pt; display:block;";
			CssWhenTop.ActiveText = "text-decoration: overline;";
			CssWhenBottom.ActiveText = "text-decoration: underline;";


			CssStyleSheetContent = "float:left;padding:1px";

			CssStyleLayoutTable = "vertical-align: top;";
            
			FCssTags = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
			FCssTags.Add("width", "100pt");
			FCssTags.Add("bordercolor", "#90bade");
			FCssTags.Add("pagecolor", "white");
			FCssTags.Add("activetabfg", "black");
			FCssTags.Add("unselectedtabbg", "#F0F0F0");
			FCssTags.Add("unselectedtabfg", "silver");
			FCssTags.Add("hoverbg", "#FFFF90");
			FCssTags.Add("hoverfg", "navy");

		}
		#endregion

		#region Events
		/// <summary>
		/// Use this event to customize the text and link on the individual tabs. The tab style itself must be modified with the CSS properties.
		/// </summary>
		[Category("Sheet Selector"),
		Description("Use this event to customize the text and link on the individual tabs.")]      
		public event SheetSelectorEntryEventHandler SheetSelectorEntry;

		/// <summary>
		/// Replace this event when creating a custom descendant of TStandardSheetSelector.
		/// </summary>
		/// <param name="e"></param>
		protected virtual void OnSheetSelectorEntry(SheetSelectorEntryEventArgs e)
		{
			if (SheetSelectorEntry!=null) SheetSelectorEntry(this, e);
		}

		#endregion

		#region Properties
		/// <summary>
		/// Style to be applied to the selector. This is a general setting, you can later further customize the style when
		/// the selector is at the Left, Top, Right or Bottom with the corresponding <see cref="CssWhenLeft"/>, <see cref="CssWhenTop"/>, <see cref="CssWhenRight"/> and <see cref="CssWhenBottom"/> properties.
		/// </summary>
		public TStandardSheetSelectorStyles CssGeneral {get {return FCssGeneral;}}

		/// <summary>
		/// Specific style to be applied to the selector when it goes at the left. This style will override the style you specify with <see cref="CssGeneral"/>
		/// </summary>
		public TStandardSheetSelectorStyles CssWhenLeft {get {return FCssWhenLeft;}}

		/// <summary>
		/// Specific style to be applied to the selector when it goes at the top. This style will override the style you specify with <see cref="CssGeneral"/>
		/// </summary>
		public TStandardSheetSelectorStyles CssWhenTop {get {return FCssWhenTop;}}
		
		/// <summary>
		/// Specific style to be applied to the selector when it goes at the right. This style will override the style you specify with <see cref="CssGeneral"/>
		/// </summary>
		public TStandardSheetSelectorStyles CssWhenRight {get {return FCssWhenRight;}}
		
		/// <summary>
		/// Specific style to be applied to the selector when it goes at the bottom. This style will override the style you specify with <see cref="CssGeneral"/>
		/// </summary>
		public TStandardSheetSelectorStyles CssWhenBottom {get {return FCssWhenBottom;}}

        /// <summary>
        /// Style to be applied to the sheet content.
        /// </summary>
        public string CssStyleSheetContent { get { return FCssStyleSheetContent; } set { FCssStyleSheetContent = value; } }

		/// <summary>
		/// Style to be applied to the layout table if <see cref="LayoutTable"/> is true.
		/// </summary>
		public string CssStyleLayoutTable { get { return FCssStyleLayoutTable; } set { FCssStyleLayoutTable = value; } }
		

		/// <summary>
		/// This property has a list of Macros that you can use in the CSS definitions. You can reference this value in the
		/// CSS properties by using &lt;#variable&gt;
		/// <p>For example, you could set a Macro "Mycolor" with CssTags.Add("mycolor", "red"); and then
		/// define a CssProperty: CssWhenTop.Main = "background-color:&lt;#mycolor&gt;";
		/// </p>
		/// <p>This method by default contains the following Macros:
		/// <list type="bullet">
		/// <item>width: Width of the selector when it is on the right or on the left.</item>
		/// <item>bordercolor: Color for the borders.</item> 
		/// <item>pagecolor: Color of the page. This is normally white, and will be used in the color of the
		/// active tab, so it blends with the page background.</item>
		/// <item>activetabfg: Color for the text in the active tab. The default is black.</item>
		/// <item>unselectedtabbg: Color for the background of the unselected tabs. The default is gray, so the active tab stands out.</item>
		/// <item>unselectedtabfg: Color for the text in the unselected tabs.</item>
		/// <item>hoverbg: Color for the unselected tab when you hover the mouse over it.</item>
		/// <item>hoverfg: Color for the text of the unselected tab when you hover the mouse over it.</item>
		/// </list>
		/// </p>
		/// <p>You can modify those Macros or add your own definitions here and use them when defining your CSS.</p>
		/// <b>Note that the variables are not case sensitive. You can write them in any combination of lowercase and uppercase.</b>
		/// </summary>
		public Dictionary<string, string> CssTags {get {return FCssTags;}}

        /// <summary>
        /// When this property is true (the default), both selectors at the left and at the right will be layed out in a table.
        /// This has the advantage that block will not wrap down when resizing the window. But if you would prefer not to use tables for layout,
        /// you can turn this property off, and the layout will be pure CSS (and it will wrap down when there is not enough space). 
        /// If you do not have a selector in the left or the right this property does nothing.
        /// </summary>
        public bool LayoutTable { get { return FLayoutTable; } set { FLayoutTable = value; } }

		/// <summary>
		/// When true (the default) and the sheets have a tab color defined in Excel, FlexCel will use this color to render the sheet tabs.
		/// If false, the default tab color will be used. Note that if the sheets don't have a color defined in Excel, also the default tab color will be used.
		/// </summary>
		public bool UseSheetTabColors {get {return FUseSheetTabColors;} set { FUseSheetTabColors = value; } }

		#endregion

		#region Write Sheet Selector

        private static string GetClass(string Selector, string TabPos)
		{
			return String.Format("class = '{0} {1}'", Selector, Selector + "_" + TabPos);  //Last element could be just TabPos, but this will not work in IE6
		}

        /// <summary>
        /// This method is overriden to add a table for layout when <see cref="LayoutTable"/> is true.
        /// See the documentation in <see cref="TSheetSelector.BeforeDrawOneSheetSelector"/> for more information on this method.
        /// </summary>
        /// <param name="Reference"></param>
        public override void BeforeDrawOneSheetSelector(TSheetSelectorPosition Reference)
        {
            if (!LayoutTable) return;

            if (Reference == TSheetSelectorPosition.Left && (SheetSelectorPosition & (TSheetSelectorPosition.Left | TSheetSelectorPosition.Right)) != 0)
            {
                WriteLn("<table border='0' cellpadding='0' cellspacing='0' summary='Layout container'><tr><td class = 'selector_layout'>");
            }
            if (Reference == TSheetSelectorPosition.Right && (SheetSelectorPosition & TSheetSelectorPosition.Right) != 0)
            {
                WriteLn("</td ><td class = 'selector_layout'>");
            }
        }

        /// This method is overriden to add a table for layout when <see cref="LayoutTable"/> is true.
        /// See the documentation in <see cref="TSheetSelector.AfterDrawOneSheetSelector"/> for more information on this method.
        public override void AfterDrawOneSheetSelector(TSheetSelectorPosition Reference)
        {
            if (!LayoutTable) return;
            
            if (Reference == TSheetSelectorPosition.Left && (SheetSelectorPosition & TSheetSelectorPosition.Left) != 0)
            {
                WriteLn("</td><td class = 'selector_layout'>");
            }
            if (Reference == TSheetSelectorPosition.Right && (SheetSelectorPosition & (TSheetSelectorPosition.Left | TSheetSelectorPosition.Right)) != 0)
            {
                WriteLn("</td></tr></table>");
            }
        }

		private string GetBkTabColor(ExcelFile xls, bool ActiveSheet)
		{
			if (!UseSheetTabColors || ActiveSheet) return String.Empty;
           
			TExcelColor sheetColor = xls.SheetTabColor;
			if (sheetColor.IsAutomatic) return String.Empty;

			return " style = 'background:" + THtmlColors.GetColor(sheetColor.ToColor(xls)) + "'";
		}

		/// <summary>
		/// This method overrides the abstract parent to provide a CSS implementation for tabs.
		/// Override this method on a child class if you want to completely customize how the Sheet Selector is drawn.
		/// Normally you can just change the CSS properties of this class to customize the SheetSelector, but you can use this if you want to 
		/// provide a completely different selector. You can use the <see cref="FlexCel.Render.TSheetSelector.Links"/> collection to know which hyperlinks to place in each place.
		/// </summary>
		/// <param name="xls">ExcelFile we are exporting. It is positioned on the sheet we are exporting, but its active sheet can be changed on this method
		/// and there is no need to restore it. It will be restored by the framework.</param>
		/// <param name="Position">Position where the sheet selector will be placed. Note that this method will always
		/// be called with only one value, but it might be called more than once if constants in the TSheetSelectorPosition enumeration are
		/// or'ed together.</param>
		public override void DrawOneSheetSelector(ExcelFile xls, TSheetSelectorPosition Position)
		{
			string TabPos = "top";
			switch (Position)
			{
				case TSheetSelectorPosition.Left: TabPos = "left"; break;
				case TSheetSelectorPosition.Right: TabPos = "right"; break;
				case TSheetSelectorPosition.Bottom: TabPos = "bottom"; break;
			}

			WriteLn("<div " + GetClass("sheetselector", TabPos) + ">");
			WriteLn("<ul " + GetClass("sheetselector_list", TabPos) + ">");
			int SaveActiveSheet = xls.ActiveSheet;
			try
			{
				int RealSheet = 1;
				for (int s = 1; s<= xls.SheetCount; s++)
				{
					xls.ActiveSheet = s;
					if (xls.SheetVisible != TXlsSheetVisible.Visible) continue;
					string TabStyle = s == SaveActiveSheet? "sheetselector_active_tab" : "sheetselector_unselected_tab";
					string Link = Links[RealSheet - 1];
                    string EntryText = xls.SheetName;

                    SheetSelectorEntryEventArgs e = new SheetSelectorEntryEventArgs(xls, s, SaveActiveSheet, Link, EntryText);
                    OnSheetSelectorEntry(e);

					string LinkStr1 = s== SaveActiveSheet? "<span " + GetClass("sheetselector_active_text", TabPos) + ">": "<a href='" + e.Link + "' " + GetClass("sheetselector_link", TabPos) + ">";
					string LinkStr2 = s== SaveActiveSheet? "</span>": "</a>";
					string DefaultCell = "  <li " + GetClass(TabStyle, TabPos) +   GetBkTabColor(xls, s == SaveActiveSheet) + ">"+
                        LinkStr1 + EncodeAsHtml(e.EntryText) + LinkStr2 + "</li>";


					WriteLn (DefaultCell);

					RealSheet++;
				}
			}
			finally
			{
				xls.ActiveSheet = SaveActiveSheet;
			}
			WriteLn("</ul>");
			WriteLn("</div>");
		}

		/// <summary>
		/// This is an utility method that will replace all the &lt;#...&gt; macros inside a tag by its values.
		/// Normally you do not need to call this method, since it is called automatically by <see cref="WriteOneCssRule"/>
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		protected string ReplaceMacros(string s)
		{
			return FlexCelHtmlExport.ReplaceMacros(s, FCssTags);
		}

		/// <summary>
		/// This method is an utility to write one rule in a TStandarSheetSelectorStyles class.
		/// </summary>
		/// <param name="Selector">A string with the style for the class. It can contain Macros (&lt;#...&gt;) and they will be replaced when writing the rule.</param>
        /// <param name="Rule">Name of the rule.</param>
        /// <param name="SubSelector">A string to be added at the end of the rule name.</param>
		protected void WriteOneCssRule(string Selector, string Rule, string SubSelector)
		{
			if (Selector == null) return;
			string RealSelector = ReplaceMacros(Selector); 
			WriteLn(" .sheetselector"+ Rule + SubSelector + "{"+ RealSelector + "}");
		}

		/// <summary>
		/// This method is an utility to write all the classes of an TStandardSheetSelectorStyles class.
		/// </summary>
		/// <param name="Selector">Collection of rules to write.</param>
        /// <param name="subSelector">A string that will be added at the end of every rule name.</param>
		protected void WriteOneCssClass(TStandardSheetSelectorStyles Selector, string subSelector)
		{
			if (Selector == null) return;

			WriteOneCssRule(Selector.Main, "",  subSelector);
			WriteOneCssRule(Selector.ActiveTab, "_active_tab",  subSelector);
			WriteOneCssRule(Selector.ActiveText, "_active_text",  subSelector);
			WriteOneCssRule(Selector.Links, "_link",  subSelector);
			WriteOneCssRule(Selector.List, "_list",  subSelector);
			WriteOneCssRule(Selector.UnselectedTab, "_unselected_tab",  subSelector);
			WriteOneCssRule(Selector.UnselectedTabHover, "_unselected_tab:hover",  subSelector);

		}

		/// <summary>
		/// This method is in charge of writing the style definitions in the header of the html file.
		/// Note that you normally do not need to override this method, you can just change the CSS properties of this class.
		/// You can override this method if you want full control on how to export the classes.
		/// </summary>
		public override void WriteCssClasses()
		{
			WriteOneCssClass(CssGeneral, "");
			WriteOneCssClass(CssWhenLeft, "_left");  //sadly we cannot use WriteOneCssClass(CssWhenLeft, ".left"); because IE6 doesn't understand it.
			WriteOneCssClass(CssWhenTop, "_top");
			WriteOneCssClass(CssWhenRight, "_right");
			WriteOneCssClass(CssWhenBottom, "_bottom");
			WriteLn(" .sheet_content {" + CssStyleSheetContent + "}");
			WriteLn(" .selector_layout {" + CssStyleLayoutTable + "}");
		}
		#endregion

	}
	#endregion
}

