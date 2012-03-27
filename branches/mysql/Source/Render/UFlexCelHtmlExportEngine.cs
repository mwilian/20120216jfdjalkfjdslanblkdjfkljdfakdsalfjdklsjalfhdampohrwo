#region Using clauses
using System;
using System.Text;
using System.Globalization;
using FlexCel.Core;
using System.IO;
using System.Reflection;
using System.Security;
using System.Security.Permissions;
using System.Collections.Generic;

#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
#else
#if (WPF)
	using RectangleF = System.Windows.Rect;
	using SizeF = System.Windows.Size;
	using PointF = System.Windows.Point;
	using real = System.Double;
	
	using System.Windows.Media;
#else
using real = System.Single;
using Colors = System.Drawing.Color;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.CompilerServices;
#endif
#endif
#endregion

namespace FlexCel.Render
{
    #region Supporting classes

    /// <summary>
    /// Indicates how much of the report has been generated.
    /// </summary>
    public class FlexCelHtmlExportProgress
    {
        private volatile int FRow;
        private volatile int FTotalRows;
        private volatile int FSheetNumber;

        internal FlexCelHtmlExportProgress()
        {
            Clear(-1);
        }

        internal void Clear(int aSheetNumber)
        {
            FRow = 0;
            FTotalRows = 0;
            FSheetNumber = 1;
            FSheetNumber = aSheetNumber;
        }

        internal void IncRow()
        {
            FRow++; //no need to interlocked here, since it is only accessed by one thread and it is volatile.
        }

        internal void SetTotalRows(int value)
        {
            FTotalRows = value;
        }

        internal void SetSheetNumber(int value)
        {
            FSheetNumber = value;
        }

        /// <summary>
        /// The row that is being written.
        /// </summary>
        public int Row { get { return FRow; } }

        /// <summary>
        /// The total number of rows exporting.
        /// </summary>
        public int TotalRows { get { return FTotalRows; } }

        /// <summary>
        /// The sheet we are exporting.
        /// </summary>
        public int SheetNumber { get { return FSheetNumber; } }
    }

    /// <summary>
    /// Encapsulates the information needed to create external CSS files.
    /// Note that if you use the same TCssInformation instance to create different html files, the CSS file created will be only one.
    /// </summary>
    public class TCssInformation
    {
        #region Privates
        private TextWriter FData;
        private string FUrl;
        internal TUsedFormatsCache UsedFormats;
        #endregion

        /// <summary>
        /// Creates a new instance of TCssInformation.
        /// </summary>
        /// <param name="aData">TextWriter where an external CSS file will be stored. If null, no CSS file will be created. A link to a CSS file might be still included if you set the Url to a non null value.</param>
        /// <param name="aUrl">URL of the css file that will be linked to this file. If null, all css information will be stored inside the html file.</param>
        public TCssInformation(TextWriter aData, string aUrl)
        {
            FData = aData;
            FUrl = aUrl;
            UsedFormats = new TUsedFormatsCache();
        }

        /// <summary>
        /// TextWriter where an external CSS file will be stored. If null, no CSS file will be created. A link to a CSS file might be still included if you set the Url to a non null value.
        /// </summary>
        public TextWriter Data { get { return FData; } }

        /// <summary>
        /// URL of the css file that will be linked to this file. If null, no CSS file will be created and all CSS information will be stored inside the html file.
        /// </summary>
        public string Url { get { return FUrl; } }
    }


    /// <summary>
    /// Stores extra data to write in the HTML file.
    /// </summary>
    public class THtmlExtraInfo
    {
        #region Privates
        private string FTitle;
        private string[] FMeta;
        private string[] FHeadStart;
        private string[] FHeadEnd;
        private string[] FBodyStart;
        private string[] FBodyEnd;
        private string[] FPrintAreaSeparator;

        #endregion

        /// <summary>
        /// Title of the HTML file. If left null, the title of the page will be used.
        /// </summary>
        public string Title { get { return FTitle; } set { FTitle = value; } }

        /// <summary>
        /// Extra strings to be added in the meta section of the header. You could specify keywords here, for example.
        /// </summary>
        public string[] Meta { get { return FMeta; } set { FMeta = value; } }

        /// <summary>
        /// Extra strings to be added after the opening &lt;head&gt; tag.
        /// </summary>
        public string[] HeadStart { get { return FHeadStart; } set { FHeadStart = value; } }

        /// <summary>
        /// Extra strings to be added before the closing &lt;/head&gt; tag.
        /// </summary>
        public string[] HeadEnd { get { return FHeadEnd; } set { FHeadEnd = value; } }

        /// <summary>
        /// Extra strings to be added after the opening &lt;body&gt; tag and before the table data.
        /// </summary>
        public string[] BodyStart { get { return FBodyStart; } set { FBodyStart = value; } }

        /// <summary>
        /// Extra strings to be added before the closing &lt;/body&gt; tag and after the table data.
        /// </summary>
        public string[] BodyEnd { get { return FBodyEnd; } set { FBodyEnd = value; } }

        /// <summary>
        /// Extra strings to be added after each section of a non-contiguous print area has been exported.
        /// Note that normally print areas are square, and in that case this property has no effect. This property
        /// only works when the print area has more than one section.
        /// </summary>
        public string[] PrintAreaSeparator { get { return FPrintAreaSeparator; } set { FPrintAreaSeparator = value; } }
    }

    /// <summary>
    /// Where to place the tabs for selecting a sheet when exporting multiple sheets.
    /// You might combine more than one, for example, to have tabs at the top and bottom:
    /// C#: TSheetSelectorPosition.Top | TSheetSelectorPosition.Bottom
    /// VB.NET, Delphi.NET: TSheetSelectorPosition.Top or TSheetSelectorPosition.Bottom
    /// </summary>
    [Flags]
    public enum TSheetSelectorPosition
    {
        /// <summary>
        /// Do not add tabs for the sheets.
        /// </summary>
        None = 0,

        /// <summary>
        /// Add the tabs at the left of the html document.
        /// </summary>
        Left = 1,

        /// <summary>
        /// Add the tabs at the right of the html document.
        /// </summary>
        Right = 2,

        /// <summary>
        /// Add the tabs at the top of the html document.
        /// </summary>
        Top = 4,

        /// <summary>
        /// Add the tabs at the bottom of the html document.
        /// </summary>
        Bottom = 8

    }

    /// <summary>
    /// Possible values in which we can save an image when exporting to HTML.
    /// </summary>
    public enum THtmlImageFormat
    {
        /// <summary>
        /// Save image as PNG. Note that transparency in PNG is not supported in ie6 unless you apply a fix.
        /// </summary>
        Png,

        /// <summary>
        /// Save image as GIF. Image will be converted to 256 colors.
        /// </summary>
        Gif,

        /// <summary>
        /// Save image as JPEG. Note that transparency is not supported in JPEG.
        /// </summary>
        Jpeg
    }

    #endregion

    #region Event Handlers

    #region Image Information Event Handler
    /// <summary>
    /// Arguments passed on <see cref="FlexCel.Render.FlexCelHtmlExport.OnGetImageInformation"/>, 
    /// </summary>
    public class ImageInformationEventArgs : EventArgs
    {
        private readonly ExcelFile FWorkbook;

        private int FObjectIndex;
        private TShapeProperties FShapeProps;
        private Stream FImageStream;
        private string FImageFile;
        private string FImageLink;
        private string FAlternateText;
        private string FHyperLink;
        private THtmlImageFormat FSavedImageFormat;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        /// <param name="aWorkbook">See <see cref="Workbook"/></param>
        /// <param name="aObjectIndex">See <see cref="ObjectIndex"/></param>
        /// <param name="aShapeProps">See <see cref="ShapeProps"/></param>
        /// <param name="aImageStream">See <see cref="ImageStream"/></param>
        /// <param name="aImageFile">See <see cref="ImageFile"/></param>
        /// <param name="aImageLink">See <see cref="ImageLink"/></param>
        /// <param name="aAlternateText">See <see cref="AlternateText"/></param>
        /// <param name="aHyperLink">See <see cref="HyperLink"/></param>
        /// <param name="aSavedImageFormat">See <see cref="SavedImageFormat"/></param>
        public ImageInformationEventArgs(ExcelFile aWorkbook, int aObjectIndex, TShapeProperties aShapeProps, Stream aImageStream, string aImageFile, string aImageLink, string aAlternateText, string aHyperLink, THtmlImageFormat aSavedImageFormat)
        {
            FWorkbook = aWorkbook;
            FObjectIndex = aObjectIndex;
            FShapeProps = aShapeProps;
            FImageStream = aImageStream;
            FImageFile = aImageFile;
            FImageLink = aImageLink;
            FAlternateText = aAlternateText;
            FHyperLink = aHyperLink;
            FSavedImageFormat = aSavedImageFormat;
        }

        /// <summary>
        /// ExcelFile with the image, positioned in the sheet that we are rendering. 
        /// Make sure if you modify ActiveSheet of this instance to restore it back to the original value before exiting the event.
        /// </summary>
        public ExcelFile Workbook { get { return FWorkbook; } }

        /// <summary>
        /// Object index of the object being rendered. You can use xls.GetObject(objectIndex) to get the object properties, or you can use this
        /// property to attach an unique number in the sheet to the image filename. If the image is not an object (for example it is a rotated text)
        /// this property will be -1.
        /// </summary>
        public int ObjectIndex { get { return FObjectIndex; } }

        /// <summary>
        /// Shape properties of the object being rendered. You can use them to get the name of the object, its size, etc.
        ///  If the image is not an object (for example it is a rotated text)
        /// this property will be null. 
        /// </summary>
        public TShapeProperties ShapeProps { get { return FShapeProps; } }


        /// <summary>
        /// The stream where the images will be saved. Keep it null to store the image as a file using <see cref="ImageFile"/>.
        /// When saving as MHTML this parameter does nothing, since all images will be saved in the same MTHML stream.
        /// </summary>
        public Stream ImageStream { get { return FImageStream; } set { FImageStream = value; } }

        /// <summary>
        /// The file where the image will be saved. If <see cref="ImageStream"/> is not null, this property will do nothing.
        /// If both this property and <see cref="ImageStream"/> are null, the image will not be saved.
        /// When saving as MHTML this parameter does nothing, since all images will be saved in the same MTHML stream.
        /// </summary>
        public string ImageFile { get { return FImageFile; } set { FImageFile = value; } }


        /// <summary>
        /// The link that will be inserted in the html file. Change it if you change the default image location.
        /// Set it to null to not add a link to this image in the generated html file. If you want to avoid exporting all images,
        /// you can use <see cref="THidePrintObjects"/> for that. But if you just want to avoid exporting one image in a file, you can do it
        /// by setting <see cref="ImageStream"/>, <see cref="ImageFile"/> and this property to null.
        /// </summary>
        public string ImageLink { get { return FImageLink; } set { FImageLink = value; } }

        /// <summary>
        /// Alternate text for the image, to show in the "ALT" tag when a browser cannot display images.
        /// By default this is set to the text in the box "Alternative Text" in the web tab on the image properties.
        /// If no Alternative text is supplied in the file, the image name will be used here.
        /// </summary>
        public string AlternateText { get { return FAlternateText; } set { FAlternateText = value; } }

        /// <summary>
        /// Hyperlink where the image will point to. This is automatically read from the image hyperlink if it has one,
        /// but you can modify, delete or add a new hyperlink for any image with this property.
        /// </summary>
        public string HyperLink { get { return FHyperLink; } set { FHyperLink = value; } }

        /// <summary>
        /// File format in which to save this image. Do not modify it to keep the default format.
        /// </summary>
        public THtmlImageFormat SavedImageFormat { get { return FSavedImageFormat; } set { FSavedImageFormat = value; } }


    }

    /// <summary>
    /// Delegate used to specify where to store the images on a page.
    /// </summary>
    public delegate void ImageInformationEventHandler(object sender, ImageInformationEventArgs e);
    #endregion

    #region Save Image Event Handler
    /// <summary>
    /// Arguments passed on <see cref="FlexCel.Render.FlexCelHtmlExport.OnSaveImage"/>, 
    /// </summary>
    public class SaveImageEventArgs : EventArgs
    {
        private readonly ExcelFile FWorkbook;

        private readonly int FObjectIndex;
        private readonly TShapeProperties FShapeProps;
        private readonly string FImageFile;
        private readonly string FImageLink;
        private readonly string FAlternateText;
        private readonly THtmlImageFormat FSavedImageFormat;

        private Image FImageToSave;
        private bool FProcessed;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        /// <param name="aWorkbook">See <see cref="Workbook"/></param>
        /// <param name="aObjectIndex">See <see cref="ObjectIndex"/></param>
        /// <param name="aShapeProps">See <see cref="ShapeProps"/></param>
        /// <param name="aImageFile">See <see cref="ImageFile"/></param>
        /// <param name="aImageLink">See <see cref="ImageLink"/></param>
        /// <param name="aAlternateText">See <see cref="AlternateText"/></param>
        /// <param name="aSavedImageFormat">See <see cref="SavedImageFormat"/></param>
        /// <param name="aImageToSave">See <see cref="ImageToSave"/></param>
        public SaveImageEventArgs(ExcelFile aWorkbook, int aObjectIndex, TShapeProperties aShapeProps, string aImageFile, string aImageLink, string aAlternateText, THtmlImageFormat aSavedImageFormat, Image aImageToSave)
        {
            FWorkbook = aWorkbook;
            FObjectIndex = aObjectIndex;
            FShapeProps = aShapeProps;
            FImageFile = aImageFile;
            FImageLink = aImageLink;
            FAlternateText = aAlternateText;
            FSavedImageFormat = aSavedImageFormat;
            FImageToSave = aImageToSave;
            FProcessed = false;
        }

        /// <summary>
        /// ExcelFile with the image, positioned in the sheet that we are rendering. 
        /// Make sure if you modify ActiveSheet of this instance to restore it back to the original value before exiting the event.
        /// </summary>
        public ExcelFile Workbook { get { return FWorkbook; } }

        /// <summary>
        /// Object index of the object being rendered. You can use xls.GetObject(objectIndex) to get the object properties. 
        /// If the image is not an object (for example it is a rotated text)
        /// this property will be -1.
        /// </summary>
        public int ObjectIndex { get { return FObjectIndex; } }

        /// <summary>
        /// Shape properties of the object being rendered. You can use them to get the name of the object, its size, etc.
        ///  If the image is not an object (for example it is a rotated text)
        /// this property will be null. 
        /// </summary>
        public TShapeProperties ShapeProps { get { return FShapeProps; } }


        /// <summary>
        /// The file where the image is expected to be saved.
        /// </summary>
        public string ImageFile { get { return FImageFile; } }


        /// <summary>
        /// The link that will be inserted in the html file.
        /// </summary>
        public string ImageLink { get { return FImageLink; } }

        /// <summary>
        /// Alternate text for the image, to show in the "ALT" tag when a browser cannot display images.
        /// By default this is set to the text in the box "Alternative Text" in the web tab on the image properties.
        /// If no Alternative text is supplied in the file, the image name will be used here.
        /// </summary>
        public string AlternateText { get { return FAlternateText; } }

        /// <summary>
        /// File format in which the image is. 
        /// </summary>
        public THtmlImageFormat SavedImageFormat { get { return FSavedImageFormat; } }

        /// <summary>
        /// Image that will be saved. You can use it to save it yourself.
        /// </summary>
        public Image ImageToSave { get { return FImageToSave; } }

        /// <summary>
        /// Set this property to true if you have taken care of saving the image, and FlexCel does not need to save it.
        /// If you just used this event to get information on the image being saved, but would like to keep the normal flux,
        /// set it to false.
        /// </summary>
        public bool Processed { get { return FProcessed; } set { FProcessed = value; } }
    }

    /// <summary>
    /// Delegate used to specify where to store the images on a page.
    /// </summary>
    public delegate void SaveImageEventHandler(object sender, SaveImageEventArgs e);
    #endregion

    #region Named Range Event Handler
    /// <summary>
    /// Arguments passed on <see cref="FlexCel.Render.FlexCelHtmlExport.OnNamedRangeExport"/>, 
    /// </summary>
    public class NamedRangeExportEventArgs : EventArgs
    {
        private readonly ExcelFile FWorkbook;

        private readonly int FSheet;
        private readonly int FRow;
        private readonly int FCol;
        private readonly TXlsNamedRange FNamedRange;

        private string FNameId;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        public NamedRangeExportEventArgs(ExcelFile aWorkbook, int aSheet, int aRow, int aCol, TXlsNamedRange aNamedRange, string aNameId)
        {
            FWorkbook = aWorkbook;
            FSheet = aSheet;
            FRow = aRow;
            FCol = aCol;
            FNamedRange = aNamedRange;
            FNameId = aNameId;
        }

        /// <summary>
        /// ExcelFile with the name, positioned in the sheet that we are rendering. 
        /// Make sure if you modify ActiveSheet of this instance to restore it back to the original value before exiting the event.
        /// </summary>
        public ExcelFile Workbook { get { return FWorkbook; } }

        /// <summary>
        /// Sheet index (1 based) of the html cell we are exporting.
        /// </summary>
        public int Sheet { get { return FSheet; } }

        /// <summary>
        /// Row index (1 based) of the html cell we are exporting. This number should be the same as the first row in the <see cref="NamedRange"/>.
        /// </summary>
        public int Row { get { return FRow; } }

        /// <summary>
        /// Column index (1 based) of the html cell we are exporting. This number should be the same as the first column in the <see cref="NamedRange"/>.
        /// </summary>
        public int Col { get { return FCol; } }


        /// <summary>
        /// Named range that is being exported.
        /// </summary>
        public TXlsNamedRange NamedRange { get { return FNamedRange; } }

        /// <summary>
        /// This property is by default the same as <see cref="NamedRange"/>.Name.  If you want to change the id
        /// of the span that will be exported to HTML, change it to the new value. To not export this name, set it to null.
        /// </summary>
        public string NameId { get { return FNameId; } set { FNameId = value; } }
    }

    /// <summary>
    /// Delegate used to customize exporting of named ranges on a page.
    /// </summary>
    public delegate void NamedRangeExportEventHandler(object sender, NamedRangeExportEventArgs e);
    #endregion

    #region SheetSelectorEntry Event Handler
    /// <summary>
    /// Arguments passed on <see cref="FlexCel.Render.TStandardSheetSelector.OnSheetSelectorEntry"/>, 
    /// </summary>
    public class SheetSelectorEntryEventArgs : EventArgs
    {
        private readonly ExcelFile FWorkbook;
        private readonly int FActiveSheet;
        private readonly int FRenderingSheet;

        private string FLink;
        private string FEntryText;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        /// <param name="aWorkbook">See <see cref="Workbook"/></param>
        /// <param name="aActiveSheet">See <see cref="ActiveSheet"/></param>
        /// <param name="aRenderingSheet">See <see cref="RenderingSheet"/></param>
        /// <param name="aLink">See <see cref="Link"/></param>
        /// <param name="aEntryText">See <see cref="EntryText"/></param>
        public SheetSelectorEntryEventArgs(ExcelFile aWorkbook, int aActiveSheet, int aRenderingSheet, string aLink, string aEntryText)
        {
            FWorkbook = aWorkbook;
            FActiveSheet = aActiveSheet;
            FRenderingSheet = aRenderingSheet;
            FLink = aLink;
            FEntryText = aEntryText;
        }

        /// <summary>
        /// ExcelFile we are drawing the sheet selector in, positioned in the sheet that we are rendering. 
        /// Make sure if you modify ActiveSheet of this instance to restore it back to the original value before exiting the event.
        /// </summary>
        public ExcelFile Workbook { get { return FWorkbook; } }

        /// <summary>
        /// Sheet index of the entry. This is equivalent to <see cref="Workbook"/>.ActiveSheet
        /// </summary>
        public int ActiveSheet { get { return FActiveSheet; } }

        /// <summary>
        /// Sheet we are currently rendering. You can compare if (RenderingSheet == ActiveSheet)
        /// to highlight the active sheet when drawing the selector.
        /// </summary>
        public int RenderingSheet { get { return FRenderingSheet; } }

        /// <summary>
        /// Place where this entry should link to.
        /// </summary>
        public string Link { get { return FLink; } set { FLink = value; } }

        /// <summary>
        /// Text that will be written in this cell of the selector.
        /// </summary>
        public string EntryText { get { return FEntryText; } set { FEntryText = value; } }
    }

    /// <summary>
    /// Delegate used to specify how the Sheet Selector will be like.
    /// </summary>
    public delegate void SheetSelectorEntryEventHandler(object sender, SheetSelectorEntryEventArgs e);
    #endregion

    #endregion

    #region Html Engine
    internal class THtmlEngine
    {
        /* Pending:
         * header
         * exportcomments
         * 
         * onheader, onbody, oncell
         */
        //comments
        //theaders

        //cf and negative values in red.


        #region Privates
        private THtmlVersion FHtmlVersion;
        private THtmlFileFormat FHtmlFileFormat;
        private TMimeWriter MimeWriter;
        public THtmlFixes HtmlFixes;

        private real FHeadingWidth;
        private const string HeadingClass = "flxHeading";
        private const string MainTableClass = "flxmain_table";
        private string FHeadingStyle;

        private THidePrintObjects FHidePrintObjects;
        private bool ExportNamedRanges;

        private const int PointsPrecision = 2;
        private const string PointsFormat0 = "0.##";
        private const string PointsFormat = ":" + PointsFormat0 + "}pt";
        private string ClassPrefix;

        private real ColMultDisplay;
        private real RowMultDisplay;

        private bool VerticalTextAsImages;
        internal TImageInformation Images;

        internal int EngineRuns;

        private FlexCelHtmlExport Parent;

        private string PendingTable;
        private bool HasTd;
        private bool UseContentId;

        private bool MergeImagesInHTML32 = false;

        private const real CellPadding = 1.5f;



        private THtmlStyle HtmlStyle { get { return HtmlVersion == THtmlVersion.Html_32 ? THtmlStyle.Simple : THtmlStyle.Css; } }

        #endregion

        #region Constructor
        public THtmlEngine(string aClassPrefix, THtmlVersion htmlVersion, THtmlFileFormat aHtmlFileFormat, TMimeWriter aMimeWriter,
            THidePrintObjects aHidePrintObjects, THtmlFixes aHtmlFixes, bool aVerticalTextAsImages, TImageInformation ImageInfo,
            real aHeadingWidth, string aHeadingStyle, bool aUseContentId, bool aExportNamedRanges,
            FlexCelHtmlExport aParent)
        {
            FHtmlVersion = htmlVersion;
            FHtmlFileFormat = aHtmlFileFormat;
            MimeWriter = aMimeWriter;
            FHidePrintObjects = aHidePrintObjects;
            HtmlFixes = aHtmlFixes;
            ClassPrefix = aClassPrefix;
            EngineRuns = 0;
            VerticalTextAsImages = aVerticalTextAsImages;
            Images = ImageInfo;
            Parent = aParent;
            FHeadingWidth = aHeadingWidth;
            FHeadingStyle = aHeadingStyle;
            UseContentId = aUseContentId;
            ExportNamedRanges = aExportNamedRanges;
        }

        public void Init(ExcelFile xls)
        {
            ColMultDisplay = ExcelMetrics.ColMultDisplay(xls) * 100f / FlexCelRender.DispMul;
            RowMultDisplay = ExcelMetrics.RowMultDisplay(xls) * 100f / FlexCelRender.DispMul;
        }

        #endregion

        #region Properties
        public THtmlVersion HtmlVersion
        {
            get { return FHtmlVersion; }
            set { FHtmlVersion = value; }
        }
        #endregion

        #region General Utilities

        /// <summary>
        /// Very simple method to ensure filenames contain only valid characters.
        /// We will consider " " and "." invalid characters, since they could have issues in some filesystems.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string EncodeFileName(string s)
        {
            StringBuilder Result = new StringBuilder(s.Length + 10);
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] >= '0' && s[i] <= '9')
                {
                    Result.Append(s[i]);
                    continue;
                }

                if ((s[i] >= 'A' && s[i] <= 'Z') || (s[i] >= 'a' && s[i] <= 'z'))
                {
                    Result.Append(s[i]);
                    continue;
                }

                if (s[i] == '_')
                {
                    Result.Append("__");
                    continue;
                }

                if (s[i] == ' ')
                {
                    Result.Append("_x");
                    continue;
                }

                else Result.Append("_" + ((int)s[i]).ToString("X04", CultureInfo.InvariantCulture));
            }
            return Result.ToString();
        }

        private string EndOfTag
        {
            get
            {
                return THtmlEntities.EndOfTag(FHtmlVersion);
            }
        }

        #region Writers
        public void Write(TextWriter DataStream, string Text)
        {
            if (MimeWriter == null)
            {
                DataStream.Write(Text);
            }
            else
            {
                MimeWriter.WriteQuotedPrintable(DataStream, Text);
            }
        }

        public void WriteLn(TextWriter DataStream, string Text)
        {
            Write(DataStream, Text + "\r\n");  //\r\n is the standard in mime CR, so we will use it here too.
        }

        private void WriteTableTag(string Text)
        {
            PendingTable = Text;
        }

        private void WriteTableEndTag(TextWriter DataStream)
        {
            if (PendingTable == null) WriteLn(DataStream, "</table>");
        }

        private void WriteTrTag(ExcelFile xls, TextWriter DataStream, TXlsCellRange PrintRange, TExportHtmlCache Cache, string Text)
        {
            if (PendingTable != null)
            {
                WriteLn(DataStream, PendingTable);
                PendingTable = null;
                WriteColumns(xls, DataStream, PrintRange, Cache);

            }

            WriteLn(DataStream, Text);
            HasTd = false;
        }

        private void WriteTrEndTag(TextWriter DataStream)
        {
            if (!HasTd) Write(DataStream, "<td></td>"); //a tr needs a td, always or it will not validate.
            WriteLn(DataStream, "</tr>");
        }

        private void WriteTdTag(TextWriter DataStream, string Text)
        {
            HasTd = true;
            Write(DataStream, Text);
        }

        #endregion

        private string GetClass(ExcelFile xls, TFlxFormat fmt, TExportHtmlCache Cache)
        {
            TCachedFormat cfmt = UpdateCssCache(xls, Cache, fmt, Cache.SheetSelector);
            return ClassPrefix + cfmt.Id.ToString(CultureInfo.InvariantCulture);
        }

        private TCachedFormat UpdateCssCache(ExcelFile xls, TExportHtmlCache Cache, TFlxFormat fmt, TSheetSelector SheetSelector)
        {
            TCssInformation CssInfo = Cache.CssInfo;

            if (CssInfo != null && CssInfo.Url != null) //if url = null, then it is internal and has already been cached. (we cannot save "as you go" css formats when it is internal, so they have to be cached)
            {
                TCachedFormat cfmt;
                if (CssInfo.UsedFormats.TryGetValue(fmt, out cfmt)) return cfmt;

                int UniqueCount = CssInfo.UsedFormats.UniqueCount;
                CssInfo.UsedFormats.Add(xls, Cache.OnHtmlFont, fmt);
                if (CssInfo.Data != null)
                {
                    if (UniqueCount == 0)
                    {
                        WriteGenericCssClasses(CssInfo.Data, SheetSelector);
                    }

                    cfmt = CssInfo.UsedFormats[fmt];
                    if (UniqueCount < CssInfo.UsedFormats.UniqueCount) WriteOneCss(CssInfo.Data, cfmt.Fmt, cfmt.Id);
                    return cfmt;
                }
            }
            return Cache.UsedFormats[fmt];

        }

        private string GetClassString(ExcelFile xls, int xf, TExportHtmlCache Cache)
        {
            if (xf <= 0) return String.Empty;
            return GetClassString(xls, xls.GetFormat(xf), Cache);
        }

        private string GetClassString(ExcelFile xls, TFlxFormat fmt, TExportHtmlCache Cache)
        {
            if (HtmlVersion == THtmlVersion.Html_32) return String.Empty;
            return "class='" + GetClass(xls, fmt, Cache) + "'";
        }


        private real RealColWidth(ExcelFile xls, int c)
        {
            return (real)Math.Round(xls.GetColWidth(c, true) / ColMultDisplay, PointsPrecision);
        }

        private real CalcAcumColWidth(ExcelFile xls, int C1, int C2)
        {
            //We can't just have the width pre added because of rounding errors.
            real Result = 0;
            for (int i = C1; i < C2; i++) Result += RealColWidth(xls, i);
            for (int i = C1 - 1; i >= C2; i--) Result -= RealColWidth(xls, i);
            return Result;
        }

        private real RealRowHeight(ExcelFile xls, int r)
        {
            return (real)Math.Round(xls.GetRowHeight(r, true) / RowMultDisplay, PointsPrecision);
        }

        private real CalcAcumRowHeight(ExcelFile xls, int R1, int R2)
        {
            real Result = 0;
            for (int i = R1; i < R2; i++) Result += RealRowHeight(xls, i);
            for (int i = R1 - 1; i >= R2; i--) Result -= RealRowHeight(xls, i);

            return Result;
        }

        internal static string FormatStr(string s, params object[] p) //FormatStr is slow, try to avoid using it.
        {
            return String.Format(CultureInfo.InvariantCulture, s, p);
        }

        #endregion

        #region Conversion Utilities

        private string GetSpanStrLeftRight(int dir, ExcelFile xls, TExportHtmlCache Cache, int r, int cIndex, SizeF CellSize, ref int FirstCol, int RightCol, real ColWidth, out int Span, out real SpanWidth)
        {
            Span = 1;
            int cIndexSpan = 0;
            SpanWidth = ColWidth;

            int NextCol; object NextValue = null;

            //Find next unused cell.
            do
            {
                NextCol = RightCol;
                if (cIndex + cIndexSpan + 1 * dir > 0 && cIndex + cIndexSpan < xls.ColCountInRow(r))
                {
                    NextCol = xls.ColFromIndex(r, cIndex + cIndexSpan + 1 * dir) - 1 * dir;
                    int XF = -1;
                    NextValue = xls.GetCellValueIndexed(r, cIndex + cIndexSpan + 1 * dir, ref XF);
                    if (NextValue == null)
                    {
                        cIndexSpan += dir;
                    }
                }
            }
            while (NextCol * dir < RightCol * dir && NextValue == null);

            //verify there are no merged cells.
            if (dir > 0)
            {
                for (int c = FirstCol; c <= NextCol; c++)
                {
                    TXlsCellRange cr = CellMergedBounds(r, c, Cache);
                    if (cr.RowCount > 1 || cr.ColCount > 1)
                    {
                        NextCol = c - 1;
                        break;
                    }
                }
            }
            else
            {
                for (int c = FirstCol; c >= NextCol; c--)
                {
                    TXlsCellRange cr = CellMergedBounds(r, c, Cache);
                    if (cr.RowCount > 1 || cr.ColCount > 1)
                    {
                        NextCol = c + cr.ColCount - 1;
                        break;
                    }
                }
            }

            while (FirstCol * dir < NextCol * dir && CellSize.Width > ColWidth)
            {
                FirstCol += dir;
                if (xls.GetColHidden(FirstCol)) continue;
                Span++;
                ColWidth += RealColWidth(xls, FirstCol);
            }

            SpanWidth = ColWidth;

            if (Span > 1) return " colspan = '" + Span + "'";
            return String.Empty;
        }

        private string GetSpanStr(ExcelFile xls, TExportHtmlCache Cache, int r, int cIndex, SizeF CellSize, TFlxFormat fmt,
            ref int FirstCol, int LeftCol, int RightCol, real ColWidth, bool IsText, out int Span, out real SpanWidth)
        {
            Span = 1;
            SpanWidth = ColWidth;

            if (fmt.WrapText || !IsText) return String.Empty; //no span.
            if (VerticalTextAsImages && (fmt.Rotation == 90 || fmt.Rotation == 180 || fmt.Rotation == 255)) return String.Empty;

            if (fmt.HAlignment == THFlxAlignment.general || fmt.HAlignment == THFlxAlignment.left)
            {
                int Col = FirstCol; //FirstCol should not move when going left to right.
                return GetSpanStrLeftRight(1, xls, Cache, r, cIndex, CellSize, ref Col, RightCol, ColWidth, out Span, out SpanWidth);
            }

            if (fmt.HAlignment == THFlxAlignment.right)
            {
                return GetSpanStrLeftRight(-1, xls, Cache, r, cIndex, CellSize, ref FirstCol, LeftCol + 1, ColWidth, out Span, out SpanWidth);
            }

            if (fmt.HAlignment == THFlxAlignment.center) //no "center" support, but at least join cells as if it was left aligned.
            {
                int Col = FirstCol; //FirstCol should not move when going left to right.
                return GetSpanStrLeftRight(1, xls, Cache, r, cIndex, CellSize, ref Col, RightCol, ColWidth, out Span, out SpanWidth);
            }

            return String.Empty;
        }

        //Text decoration cannot be anulated from inside a text, it has to be none on the parent container.
        //So if a rich string has text decoration, the cell must have text-decoration:none;
        private static bool HasTextDecoration(ExcelFile xls, TFlxFont CellFont, TRichString CellText)
        {
            if (CellFont.Underline == TFlxUnderline.None && (CellFont.Style & TFlxFontStyles.StrikeOut) == 0) return false;

            if (CellText == null || CellText.RTFRunCount <= 0) return false;
            for (int i = 0; i < CellText.RTFRunCount; i++)
            {
                int FontIndex = CellText.RTFRun(i).FontIndex;
                if (FontIndex < 0 || FontIndex >= xls.FontCount) continue;
                TFlxFont f = xls.GetFont(FontIndex);
                if (f.Underline != CellFont.Underline || (f.Style & TFlxFontStyles.StrikeOut) != (CellFont.Style & TFlxFontStyles.StrikeOut)) return true;
            }
            return false;
        }

        private static string GetTextDecoration(ExcelFile xls, TFlxFont CellFont, TRichString CellText)
        {
            if (CellText != null && HasTextDecoration(xls, CellFont, CellText)) return String.Empty;  //When the cell has text decoration, we can not do it on the cell. We need to have spans with text decorations inside.

            string StrUnderline = CellFont.Underline != TFlxUnderline.None ? "underline" : String.Empty;

            if ((CellFont.Style & TFlxFontStyles.StrikeOut) != 0) StrUnderline += " line-through";

            if (StrUnderline.Length > 0) return "text-decoration:" + StrUnderline + ";";
            return String.Empty;
        }


        private string GetCellStyle(ExcelFile xls, TFlxFormat fmt, TFlxFormat fmt2, bool AddBorders, THAlign HAlign, bool AddIndent,
            bool AddWidth, real ColWidth, int ColSpan, bool AddHeight, real RowHeight, int RowSpan, bool AddStyleTag, TRichString CellText)
        {
            return GetCellStyle(xls, fmt, fmt2, AddBorders, HAlign, AddIndent,
                AddWidth, ColWidth, ColSpan, AddHeight, RowHeight, RowSpan, AddStyleTag, CellText, false);
        }

        private string GetCellStyle(ExcelFile xls, TFlxFormat fmt, TFlxFormat fmt2, bool AddBorders, THAlign HAlign, bool AddIndent,
            bool AddWidth, real ColWidth, int ColSpan, bool AddHeight, real RowHeight, int RowSpan, bool AddStyleTag, TRichString CellText, bool SuppressPadding)
        {
            if (HtmlVersion == THtmlVersion.Html_32) return GetCellStyle32(xls, fmt, fmt2, AddBorders, HAlign, AddIndent, SuppressPadding, fmt.WrapText || RowSpan > 1 || ColSpan > 1, ColWidth);

            StringBuilder sb = new StringBuilder();
            if (AddWidth) { sb.Append("width:"); sb.Append(ColWidth.ToString(PointsFormat0, CultureInfo.InvariantCulture)); sb.Append("pt;"); }
            if (AddHeight) { sb.Append("height:"); sb.Append(RowHeight.ToString(PointsFormat0, CultureInfo.InvariantCulture)); sb.Append("pt;"); }

            sb.Append(GetTextDecoration(xls, fmt.Font, CellText));

            if (AddBorders) GetCellBorders(sb, xls, fmt, fmt2);
            GetCellAlign(sb, fmt, HAlign);

            if (fmt.Indent > 0 && AddIndent)
            {
                real Indent = 0;

                Indent = fmt.Indent * 256f / ColMultDisplay * 1.74f;
                sb.Append("padding-left:");
                sb.Append(Indent.ToString("0.##", CultureInfo.InvariantCulture));
                sb.Append("px;");
            }
            else
            {
                if (SuppressPadding)
                {
                    sb.Append("padding:0");
                }
            }

            if (sb.Length == 0) return String.Empty;
            if (AddStyleTag) return " style ='" + sb.ToString() + "'";
            return sb.ToString();

        }

        private string GetCellStyle32(ExcelFile xls, TFlxFormat fmt, TFlxFormat fmt2, bool AddBorders, THAlign HAlign, bool AddIndent, bool SuppressPadding, bool AddWidth, real ColWidth)
        {
            StringBuilder sb = new StringBuilder();

            if (fmt.HAlignment != THFlxAlignment.general)
            {
                switch (fmt.HAlignment)
                {
                    case THFlxAlignment.center: sb.Append("align='center' "); break;
                    case THFlxAlignment.right: sb.Append("align='right' "); break;
                }
            }
            else
                switch (HAlign)
                {
                    case THAlign.Center: sb.Append("align='center' "); break;
                    case THAlign.Right: sb.Append("align='right' "); break;
                }

            switch (fmt.VAlignment)
            {
                case TVFlxAlignment.center: sb.Append("valign='middle' "); break;
                case TVFlxAlignment.top: sb.Append("valign='top' "); break; //Cells in Excel are normally aligned bottom, so we use that as default.
            }


            Color BgColor = fmt.FillPattern.Pattern == TFlxPatternStyle.Automatic ? fmt.FillPattern.BgColor.ToColor(xls) : fmt.FillPattern.FgColor.ToColor(xls);
            if ((fmt.FillPattern.Pattern == TFlxPatternStyle.Automatic || fmt.FillPattern.Pattern == TFlxPatternStyle.Solid) && BgColor.ToArgb() != Color.White.ToArgb())
            {
                sb.Append("bgcolor='" + THtmlColors.GetColor(BgColor) + "' ");
            }

            if (AddBorders)
            {
                StringBuilder bb = new StringBuilder();
                GetCellBorders(bb, xls, fmt, fmt2);
                if (bb.Length > 0)
                {
                    sb.Append("style = '" + bb.ToString() + "' "); //Not 3.2 compliant, but the only way to have borders by cell.
                }
            }

            if (AddWidth)
            {
                sb.Append("width = '");
                sb.Append((ColWidth * 96f / 72f).ToString("0.##", CultureInfo.InvariantCulture));
                sb.Append("' ");
            }


            if (sb.Length == 0) return String.Empty;
            return sb.ToString();
        }

        private void GetFontStyle32(ExcelFile xls, TFlxFormat DefaultFmt, TFlxFormat fmt, TExportHtmlCache Cache, out string Before, out string After)
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder TagsToClose = new StringBuilder();

            sb.Append(THtmlTagCreator.DiffFont(xls, DefaultFmt.Font, DefaultFmt.Font, fmt.Font, HtmlVersion, HtmlStyle, ref TagsToClose, Cache.OnHtmlFont, false));

            Before = sb.ToString();
            After = TagsToClose.ToString();
        }

        private string GetDefaultFontStyle32(ExcelFile xls, TExportHtmlCache Cache, TFlxFont DefFont)
        {
            StringBuilder sb = new StringBuilder();
            TFlxFont Font0 = new TFlxFont(); //This will be called only once. So we want to always say the typeface, be it arial or times.
            Font0.Name = String.Empty; Font0.Size20 = 0;
            Font0.Color = ColorUtil.Empty;
            THtmlTagCreator.DiffFontCss(xls, Font0, DefFont, Cache.OnHtmlFont, sb);

            return sb.ToString();

        }


        private static string GetCellAlignAndHeightTag(ExcelFile xls, TFlxFormat fmt, THAlign HAlign, real RowHeight, TRichString CellText)
        {
            StringBuilder sb = new StringBuilder();
            GetCellAlign(sb, fmt, HAlign);
            sb.Append("height:"); sb.Append(RowHeight.ToString(PointsFormat0, CultureInfo.InvariantCulture));
            sb.Append("pt;");
            sb.Append(GetTextDecoration(xls, fmt.Font, CellText));

            if (sb.Length == 0) return String.Empty;
            return " style ='" + sb.ToString() + "'";
        }

        private static void GetCellAlign(StringBuilder sb, TFlxFormat fmt, THAlign HAlign)
        {
            if (fmt.HAlignment == THFlxAlignment.general)
            {
                switch (HAlign)
                {
                    case THAlign.Center: sb.Append("text-align:center;"); break;
                    case THAlign.Right: sb.Append("text-align:right;"); break;
                }
            }
        }

        #region Borders
        private static void GetCellBorders(StringBuilder sb, ExcelFile xls, TFlxFormat fmt, TFlxFormat fmt2)
        {
            if (fmt.Borders.Left == fmt.Borders.Top && fmt.Borders.Top == fmt2.Borders.Right && fmt2.Borders.Right == fmt2.Borders.Bottom)
            {
                WriteBorder(xls, sb, fmt.Borders.Left, "border");
            }
            else
            {
                WriteBorder(xls, sb, fmt.Borders.Left, "border-left");
                WriteBorder(xls, sb, fmt.Borders.Top, "border-top");
                WriteBorder(xls, sb, fmt2.Borders.Right, "border-right");
                WriteBorder(xls, sb, fmt2.Borders.Bottom, "border-bottom");
            }
        }

        private static void WriteBorder(ExcelFile xls, StringBuilder sb, TFlxOneBorder Border, string BorderPos)
        {
            if (Border.Style == TFlxBorderStyle.None) return;
            sb.Append(BorderPos + ":");
            sb.Append(GetHtmlBorderWidth(Border.Style) + " ");
            sb.Append(GetHtmlBorderStyle(Border.Style) + " ");
            sb.Append(GetColor(xls, Border.Color, Colors.Black));
            sb.Append(";");
        }

        private static string GetHtmlBorderWidth(TFlxBorderStyle BorderStyle)
        {
            switch (BorderStyle)
            {
                case TFlxBorderStyle.Medium:
                case TFlxBorderStyle.Double:
                case TFlxBorderStyle.Medium_dashed:
                case TFlxBorderStyle.Medium_dash_dot:
                case TFlxBorderStyle.Medium_dash_dot_dot:
                    return "medium";
                case TFlxBorderStyle.Thick:
                    return "thick";
            }
            //return "thin"; //too wide in ie.
            return "1px"; //(0.5pt does not work in safari...)
        }

        private static string GetHtmlBorderStyle(TFlxBorderStyle BorderStyle)
        {
            switch (BorderStyle)
            {
                case TFlxBorderStyle.None:
                    return "none";
                case TFlxBorderStyle.Thin:
                    break;
                case TFlxBorderStyle.Medium:
                    break;
                case TFlxBorderStyle.Dashed:
                    return "dashed";
                case TFlxBorderStyle.Dotted:
                    return "dotted";
                case TFlxBorderStyle.Thick:
                    break;
                case TFlxBorderStyle.Double:
                    return "double";
                case TFlxBorderStyle.Hair:
                    break;
                case TFlxBorderStyle.Medium_dashed:
                    return "dashed";
                case TFlxBorderStyle.Dash_dot:
                    return "dashed";
                case TFlxBorderStyle.Medium_dash_dot:
                    return "dashed";
                case TFlxBorderStyle.Dash_dot_dot:
                    return "dashed";
                case TFlxBorderStyle.Medium_dash_dot_dot:
                    return "dashed";
                case TFlxBorderStyle.Slanted_dash_dot:
                    return "dashed";
            }
            return "solid";
        }

        #endregion

        internal static string GetColor(IFlexCelPalette xls, TExcelColor aColor, Color BackColor)
        {
            return THtmlColors.GetColor(aColor.ToColor(xls, BackColor));
        }

        private string EncodeAsHtml(TextWriter html, string s, TEnterStyle EnterStyle) //Note that a single quote does not get encoded by this method. So, when using it in attributes, use doble quotes.
        {
            return THtmlEntities.EncodeAsHtml(s, HtmlVersion, html.Encoding, EnterStyle);
        }

        private string EncodeHyperlink(TextWriter html, string s) //Note that a single quote does not get encoded by this method. So, when using it in attributes, use doble quotes.
        {
            if (Parent.BaseUrl != null && Parent.BaseUrl.Length > 0 && s.ToUpper(CultureInfo.InvariantCulture).StartsWith(Parent.BaseUrl.ToUpper(CultureInfo.InvariantCulture)))
            {
                s = s.Remove(0, Parent.BaseUrl.Length);
            }
            return THtmlEntities.EncodeAsHtml(s, HtmlVersion, html.Encoding, TEnterStyle.Ignore);
        }

        private string GetTitle(TextWriter html, ExcelFile xls, THtmlExtraInfo metaInfo)
        {
            string Result;
            if (metaInfo == null || metaInfo.Title == null)
            {
                if (xls == null) Result = String.Empty; else Result = xls.SheetName;
            }
            else
            {
                Result = metaInfo.Title;
            }

            return EncodeAsHtml(html, Result, TEnterStyle.Char10);
        }
        #endregion

        #region Write HTML

        internal void DumpStrings(TextWriter html, string[] info)
        {
            if (info == null) return;
            foreach (string s in info)
            {
                WriteLn(html, s);
            }

        }

        /// <summary>
        /// Writes the header of an HTML file.
        /// </summary>
        /// <param name="xls">Might be null. If null the pagetitle will be the one in metainfo</param>
        /// <param name="CssInfo"></param>
        /// <param name="hasImages"></param>
        /// <param name="html"></param>
        /// <param name="extraInfo"></param>
        /// <param name="SheetSelector"></param>
        internal void WriteHeader(ExcelFile xls, TextWriter html, THtmlExtraInfo extraInfo, TCssInformation CssInfo, bool hasImages, TSheetSelector SheetSelector)
        {
            if (HtmlVersion == THtmlVersion.Html_401)
            {
                WriteLn(html, "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01//EN\" \"http://www.w3.org/TR/html4/strict.dtd\">");

                if (HtmlFixes.Outlook2007CssSupport && hasImages && ((FHidePrintObjects & THidePrintObjects.Images) == 0))
                {
                    WriteLn(html, "<!--[if gte vml 1]><html xmlns:v=\"urn:schemas-microsoft-com:vml\"><![endif]-->");
                }

                WriteLn(html, "<html>");
            }
            else
                if (HtmlVersion == THtmlVersion.Html_32)
                {
                    WriteLn(html, "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 3.2 Final//EN\">");
                    WriteLn(html, "<html>");
                }

                else
                {
                    WriteLn(html, "<?xml version=\"1.0\" encoding=\"" + html.Encoding.WebName + "\"?>");
                    WriteLn(html, "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Strict//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd\">");
                    WriteLn(html, "<html xmlns=\"http://www.w3.org/1999/xhtml\" xml:lang=\"en\" lang=\"en\">");
                }
            WriteLn(html, "<head>");

            if (extraInfo != null) DumpStrings(html, extraInfo.HeadStart);

            if (HtmlVersion != THtmlVersion.XHTML_10)
            {
                WriteLn(html, "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=" + html.Encoding.WebName + "\"" + EndOfTag);  //First so we state the encoding.
            }
            else
            {
                WriteLn(html, "<meta http-equiv=\"Content-Type\" content=\"application/xhtml+xml; charset=" + html.Encoding.WebName + "\"" + EndOfTag);  //First so we state the encoding.
            }

            if (HtmlVersion != THtmlVersion.Html_32)
            {
                WriteLn(html, "<meta http-equiv=\"Content-Style-Type\" content=\"text/css\"" + EndOfTag);
            }

            if (extraInfo != null) DumpStrings(html, extraInfo.Meta);

            string FlexCelVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();

            WriteLn(html, "<meta name=\"Generator\" content=\"FlexCel " + FlexCelVersion + "\"" + EndOfTag);
            WriteLn(html, "<title>");
            WriteLn(html, GetTitle(html, xls, extraInfo));
            WriteLn(html, "</title>");

            WriteCss(html, CssInfo, SheetSelector);
            if (HtmlFixes.IE6TransparentPngSupport) AddIe6TransparentPngFix(html, hasImages);
            if (extraInfo != null) DumpStrings(html, extraInfo.HeadEnd);

            WriteLn(html, "</head>");
        }

        public void AddIe6TransparentPngFix(TextWriter html, bool hasImages)
        {
            if (HtmlVersion == THtmlVersion.Html_32) return;
            if (!hasImages) return;

            WriteLn(html, "<!--[if gte IE 5.5]>");
            WriteLn(html, "<![if lt IE 7]>");
            WriteLn(html, "<style type=\"text/css\">");

            WriteLn(html, "  .imagediv  {display: none}");

            WriteLn(html, "</style>");
            WriteLn(html, "<![endif]>");
            WriteLn(html, "<![endif]-->");
        }

        private real CalcTableWidth(ExcelFile xls, TXlsCellRange PrintRange)
        {
            real Gutter = xls.PrintHeadings ? FHeadingWidth : 0;
            return Gutter + CalcAcumColWidth(xls, PrintRange.Left, PrintRange.Right + 1);
        }

        internal void StartBody(TextWriter html, THtmlExtraInfo extraInfo)
        {
            WriteLn(html, "<body>");
            if (extraInfo != null) DumpStrings(html, extraInfo.BodyStart);
        }


        internal void WriteBody(ExcelFile xls, TextWriter html, TXlsCellRange[] PrintRanges, TExportHtmlCache Cache,
            FlexCelHtmlExportProgress Progress, string HtmlFileName, string RelativeImagePath, bool SaveImages, THtmlExtraInfo ExtraInfo)
        {

            if (xls.SheetType == TSheetType.Chart)
            {
                THtmlImageCache[,][] ImagesInCell;
                bool HasImages = Cache.Images.TryGetValue(1, 1, 1, 1, out ImagesInCell);

                if (HasImages)
                {
                    THtmlImageCache Img = ImagesInCell[0, 0][0];

                    Write(html, "<div>");
                    string ImgLink = "  <img src='{0}' width='{1}' height='{2}' alt=\"{3}\" ";
                    if (HtmlVersion == THtmlVersion.Html_32)
                    {
                        WriteLn(html, FormatStr(ImgLink + "border='0' {4}",
                            Img.Url, Img.SizePixels.Width, Img.SizePixels.Height, EncodeAsHtml(html, Img.AltText, TEnterStyle.Char10),
                            EndOfTag));
                    }
                    else
                    {
                        WriteLn(html, FormatStr(ImgLink +
                            "style = '" +
                            "width:{4" + PointsFormat + ";height:{5" + PointsFormat + ";' {6}",
                            Img.Url, Img.SizePixels.Width, Img.SizePixels.Height, EncodeAsHtml(html, Img.AltText, TEnterStyle.Char10),
                            Img.Dimensions.Width, Img.Dimensions.Height,
                            EndOfTag));
                    }
                    Write(html, "</div>");
                }

                return;
            }

            bool First = true;
            foreach (TXlsCellRange PrintRange in PrintRanges)
            {
                if (First) First = false;
                else
                {
                    if (ExtraInfo != null) DumpStrings(html, ExtraInfo.PrintAreaSeparator);
                }
                real TableWidth = CalcTableWidth(xls, PrintRange);

                if ((FHidePrintObjects & THidePrintObjects.Headers) == 0)
                {
                    WriteHeaders(xls, html, HtmlFileName, RelativeImagePath, Cache, SaveImages, TableWidth);
                }

                int BorderWidth = xls.PrintGridLines ? 1 : 0;
                if (HtmlVersion == THtmlVersion.Html_32)
                {
                    WriteTableTag(FormatStr("<table border='{1}' cellpadding='0' cellspacing='0' width='{0}' style='border-collapse:collapse;{2}'>", TableWidth * 96f / 72f, BorderWidth, GetDefaultFontStyle32(xls, Cache, xls.GetDefaultFont)));
                }
                else
                {
                    string Border = xls.PrintGridLines ? "border:1px solid silver;" : String.Empty;
                    WriteTableTag(FormatStr("<table class='{4}' border='{2}' cellpadding='0' cellspacing='0' style='width:{0" + PointsFormat + "; table-layout:fixed;" +
                        "border-collapse:collapse;{1}' summary=\"Excel Sheet: {3}\">", TableWidth, Border, BorderWidth, EncodeAsHtml(html, xls.SheetName, TEnterStyle.Ignore), MainTableClass));
                }

                WriteTable(xls, html, PrintRange, Cache, Progress, HtmlFileName, RelativeImagePath, SaveImages);

                WriteTableEndTag(html);

                if ((FHidePrintObjects & THidePrintObjects.Headers) == 0)
                {
                    WriteFooters(xls, html, HtmlFileName, RelativeImagePath, Cache, SaveImages, TableWidth);
                }
            }
        }

        internal void EndBody(TextWriter html, THtmlExtraInfo ExtraInfo)
        {
            if (ExtraInfo != null) DumpStrings(html, ExtraInfo.BodyEnd);
            WriteLn(html, "</body>");
        }

        private void WriteColumns(ExcelFile xls, TextWriter html, TXlsCellRange PrintRange, TExportHtmlCache Cache)
        {
            WriteColTags(xls, html, PrintRange, Cache);
            bool h32 = HtmlVersion == THtmlVersion.Html_32;

            //Outlook does not care about col tags, needs widths in the first row of cells. So we will add an empty first row. (we cannot add them in the first real row since columns might be merged)
            //On the other side, ie with table-layout:fixed will use this row to get the column widths.
            //As our td style has a padding to the left and right, and ie adds this pad to the cell width (different to all other browsers), we need
            //to add this line always, setting padding to 0.
            if (PrintRange.Right > PrintRange.Left)
            {

                if (h32) WriteLn(html, "  <tr height='0'>"); else WriteLn(html, "  <tr style='display:none'>");
                bool HasCols = false;

                if (xls.PrintHeadings)
                {
                    HasCols = true;
                    if (h32)
                    {
                        WriteLn(html, FormatStr("    <td width='{0:0.##}'></td>", FHeadingWidth * 96f / 72f));
                    }
                    else
                    {
                        WriteLn(html, FormatStr("    <td style = 'padding:0;width:{0" + PointsFormat + ";" + "'></td>", FHeadingWidth));
                    }

                }

                for (int c = PrintRange.Left; c <= PrintRange.Right; c++)
                {
                    if (xls.GetColHidden(c)) continue;
                    HasCols = true;
                    real w = RealColWidth(xls, c);
                    if (h32)
                    {
                        WriteLn(html, FormatStr("    <td width='{0:0.##}'></td>", w * 96f / 72f));
                    }
                    else
                    {
                        WriteLn(html, FormatStr("    <td style = 'padding:0;width:{0" + PointsFormat + ";" + "'></td>", w));
                    }
                }
                if (!HasCols) WriteLn(html, "    <td></td>"); //this will happen only if there are only hidden columns in the range. A tr with no td will not validate.
                WriteLn(html, "  </tr>");
            }

            if (xls.PrintHeadings)
            {
                WriteLn(html, "  <tr " + HeadingClassDef + ">");
                WriteLn(html, "  <td " + HeadingClassDef + "></td>"); //Empty box at the top-left coords.
                for (int c = PrintRange.Left; c <= PrintRange.Right; c++)
                {
                    if (xls.GetColHidden(c)) continue;

                    string ColStr = xls.OptionsR1C1 ? c.ToString(CultureInfo.CurrentCulture) : TCellAddress.EncodeColumn(c);
                    WriteLn(html, "    <td " + HeadingClassDef + ">" + ColStr + "</td>");
                }
                WriteLn(html, "  </tr>");
            }

        }

        private void WriteColTags(ExcelFile xls, TextWriter html, TXlsCellRange PrintRange, TExportHtmlCache Cache)
        {
            if (HtmlVersion == THtmlVersion.Html_32) return;

            if (xls.PrintHeadings)
            {
                WriteLn(html, FormatStr("    <col class = '{0}' style = 'width:{1" + PointsFormat + ";" + "'" + EndOfTag, HeadingClass, FHeadingWidth));
            }

            real LastW = -1;
            int LastXF = -1;
            int Span = 0;
            for (int c = PrintRange.Left; c <= PrintRange.Right + 1; c++)
            {
                int xf = -1;
                real w = -1;
                if (c <= PrintRange.Right)
                {
                    if (xls.GetColHidden(c)) continue;
                    w = RealColWidth(xls, c);
                    if (w <= 0) continue;  //Not exported.

                    xf = xls.GetColFormat(c);
                    if (xf <= 0) xf = xls.DefaultFormatId;

                    if (LastW < 0 || (LastW == w && xf == LastXF))  //We cannot compare TCachedFormats here, since they do not compare all things (for example borders).
                    {
                        Span++;
                        LastW = w;
                        LastXF = xf;
                        continue;
                    }
                }

                if (LastXF >= 0)
                {
                    string SpanStr = Span <= 1 ? String.Empty : " span='" + Span.ToString(CultureInfo.InvariantCulture) + "'";
                    TFlxFormat LastFmt = xls.GetFormat(LastXF);
                    WriteLn(html, " <col " +
                        GetClassString(xls, LastFmt, Cache) + SpanStr +
                        GetCellStyle(xls, LastFmt, LastFmt, true, THAlign.Left, false, true, LastW, Span, false, 0, 1, true, null) + EndOfTag);

                }
                Span = 1;
                LastW = w;
                LastXF = xf;
            }
        }

        private string HeadingClassDef
        {
            get
            {
                return HtmlVersion == THtmlVersion.Html_32 ? "bgcolor='#E7E7E7' align='center' style='border:1px solid black;'" : "class ='" + HeadingClass + "'";
            }
        }
        private void WriteTable(ExcelFile xls, TextWriter html, TXlsCellRange PrintRange, TExportHtmlCache Cache, FlexCelHtmlExportProgress Progress, string HtmlFileName, string RelativeImagePath, bool SaveImages)
        {
            TGraphicCanvas gr = new TGraphicCanvas();
            try
            {
                if ((FHidePrintObjects & THidePrintObjects.Hyperlynks) == 0) Cache.Hyperlinks.CacheHyperLinks(xls, PrintRange);
                if (ExportNamedRanges)
                {
                    Cache.Names = new TNamedRangeCache();
                    Cache.Names.CacheNames(xls, PrintRange);
                }
                TFlxFormat DefaultFormat = xls.GetDefaultFormat;

                bool IgnoreFormulaText = xls.IgnoreFormulaText;
                TReferenceStyle SaveRefStyle = xls.FormulaReferenceStyle;
                try
                {
                    if (xls.OptionsR1C1) xls.FormulaReferenceStyle = TReferenceStyle.R1C1; else xls.FormulaReferenceStyle = TReferenceStyle.A1;
                    xls.IgnoreFormulaText = !xls.ShowFormulaText;

                    for (int r = PrintRange.Top; r <= PrintRange.Bottom; r++)
                    {
                        if (Parent.Canceled) return;
                        Progress.IncRow();

                        if (xls.GetRowHidden(r)) continue;

                        int rh = xls.GetRowHeight(r);
                        if (rh <= 0) continue;  //Not exported.

                        int rxf = xls.GetRowFormat(r);

                        real RowHeight = RealRowHeight(xls, r);
                        bool RowAutoHeight = xls.GetAutoRowHeight(r);

                        string PageBreak = String.Empty;
                        if ((FHidePrintObjects & THidePrintObjects.PageBreaks) == 0)
                        {
                            if (xls.HasHPageBreak(r)) PageBreak = "page-break-after: always;";
                        }

                        if (HtmlVersion == THtmlVersion.Html_32)
                        {
                            WriteTrTag(xls, html, PrintRange, Cache,
                                " <tr valign='bottom' height='" + (RowHeight * 96f / 72f).ToString("0.##", CultureInfo.InvariantCulture) + "'>");
                        }
                        else
                        {
                            WriteTrTag(xls, html, PrintRange, Cache,
                                " <tr " + GetClassString(xls, rxf, Cache) + " style='height:" + RowHeight.ToString(PointsFormat0, CultureInfo.InvariantCulture)
                                + "pt;" + PageBreak + "'>");
                        }

                        if (xls.PrintHeadings)
                        {
                            WriteTdTag(html, "    <td " + HeadingClassDef + ">" + r.ToString(CultureInfo.CurrentCulture) + "</td>");
                        }

                        int ColCount = xls.ColCountInRow(r);
                        int LastUsedCol = PrintRange.Left - 1;
                        int cIndex = 1;
                        while (cIndex <= ColCount)
                        {
                            int c = xls.ColFromIndex(r, cIndex);
                            if (c > PrintRange.Right) break;

                            WriteCell(xls, html, PrintRange, gr,
                                r, rh, RowHeight, RowAutoHeight, ref LastUsedCol, ref cIndex, c, Cache, HtmlFileName, RelativeImagePath, SaveImages, DefaultFormat);
                            cIndex++;

                        }
                        FillEmptyCells(xls, html, r, LastUsedCol, PrintRange.Right + 1, Cache);
                        WriteTrEndTag(html);
                    }
                }
                finally
                {
                    xls.IgnoreFormulaText = IgnoreFormulaText;
                    xls.FormulaReferenceStyle = SaveRefStyle;
                }
            }
            finally
            {
                gr.Dispose();
            }
        }

        private void WriteCell(ExcelFile xls, TextWriter html, TXlsCellRange PrintRange, TGraphicCanvas gr,
            int r, int rh, real RowHeight, bool RowAutoHeight, ref int LastUsedCol, ref int cIndex, int c, TExportHtmlCache Cache,
            string HtmlFileName, string RelativeImagePath, bool SaveImages, TFlxFormat DefaultFormat)
        {
            if (c < PrintRange.Left) return;
            if (xls.GetColHidden(c)) return;
            int w = xls.GetColWidth(c);
            if (w <= 0) return;  //Not exported.


            TXlsCellRange MergedRange = CellMergedBounds(r, c, Cache);
            if (r > MergedRange.Top || c > MergedRange.Left)
            {
                if (c >= MergedRange.Left) FillEmptyCells(xls, html, r, LastUsedCol, c, Cache);
                LastUsedCol = MergedRange.Right;
                cIndex = xls.ColToIndex(r, LastUsedCol + 1) - 1;
                return;
            }

            int xf = -1;
            object oData = xls.GetCellValueIndexed(r, cIndex, ref xf);
            if (oData == null && MergedRange.ColCount == 1 && MergedRange.RowCount == 1) return;

            if ((xf < 0) || (xf > xls.FormatCount)) xf = xls.DefaultFormatId;

            TFormula fmla = oData as TFormula;
            if (fmla != null)
            {
                if (xls.ShowFormulaText)
                {
                    oData = fmla.Text;
                }
                else oData = fmla.Result;
            }

            Color aColor = ColorUtil.Empty;
            TFlxFormat fmt = Cache.Formats.GetCellVisibleFormatDef(xls, r, c, false);
            bool HasDate, HasTime; TAdaptativeFormats AdaptativeFormats;
            TRichString Data = TFlxNumberFormat.FormatValue(oData, fmt.Format, ref aColor, xls, out HasDate, out HasTime, out AdaptativeFormats);

            THAlign HAlign = THAlign.Left;
            bool IsText = false;
            double ActualValue = 0;
            FlexCelRender.GetDataAlign(oData, fmt, ref IsText, ref ActualValue, ref HAlign);

            real ColWidth = RealColWidth(xls, c);
            real ColSpanWidth = CalcAcumColWidth(xls, MergedRange.Left, MergedRange.Right + 1);
            Font aFont = gr.FontCache.GetFont(fmt.Font, 1);
            real md;

            real Indent = 0;
            if (fmt.Indent > 0)
            {
                Indent = fmt.Indent * 256f / ColMultDisplay * 1.74f;
            }

            if (AdaptativeFormats != null && AdaptativeFormats.WildcardPos >= 0)
            {
                TXRichString rx = new TXRichString(Data, false, 0, 0, AdaptativeFormats);
                SizeF Tx = new SizeF();
                TextPainter.AddWilcardtoLine(gr.Canvas, gr.FontCache, 1, aFont, ColSpanWidth - Indent - CellPadding * 2, ref Tx, rx);
                Data = rx.s;
            }

            SizeF CellSize = RenderMetrics.CalcTextExtent(gr.Canvas, gr.FontCache, 1, aFont, Data, out md);
            CellSize.Width += Indent + CellPadding * 2;

            if (!IsText)
            {
                if (CellSize.Width > ColSpanWidth)
                {
                    TXRichString rs = FlexCelRender.TryToFit(gr.Canvas, gr.FontCache, 1, aFont, ActualValue, ColSpanWidth, CellPadding);
                    Data = rs.s;
                    CellSize = RenderMetrics.CalcTextExtent(gr.Canvas, gr.FontCache, 1, aFont, Data, out md);
                    CellSize.Width += Indent + CellPadding * 2;
                }
            }

            int ColSpan;
            int RowSpan = 1;
            real RowSpanHeight = RowHeight;
            string SpanStr;
            int cStart = c;
            TFlxFormat fmt2 = fmt;
            if (MergedRange.ColCount > 1 || MergedRange.RowCount > 1)
            {
                SpanStr = FormatStr(" colspan = '{0}' rowspan ='{1}'", MergedRange.ColCount, MergedRange.RowCount);
                RowSpan = MergedRange.RowCount;
                ColSpan = MergedRange.ColCount;
                RowSpanHeight = CalcAcumRowHeight(xls, MergedRange.Top, MergedRange.Bottom + 1);
                fmt2 = Cache.Formats.GetCellVisibleFormatDef(xls, MergedRange.Bottom, MergedRange.Right, true);
            }
            else
            {
                SpanStr = GetSpanStr(xls, Cache, r, cIndex, CellSize, fmt, ref cStart, LastUsedCol, PrintRange.Right, ColWidth, IsText, out ColSpan, out ColSpanWidth);
                if (ColSpan > 1) fmt2 = Cache.Formats.GetCellVisibleFormatDef(xls, r, cStart + ColSpan - 1, true);
            }

            FillEmptyCells(xls, html, r, LastUsedCol, cStart, Cache);

            bool CanWrap = IsText && TCachedFormat.Wraps(fmt);

            if ((HtmlFixes.Outlook2007CssSupport || HtmlFixes.WordWrapSupport || HtmlVersion == THtmlVersion.Html_32) && !CanWrap && Data.Length > 0)
            {
                int fit = Data.Length;

                real cw = CellSize.Width - 2 * CellPadding; //We need the real size of the string here.
                RenderMetrics.FitOneLine(gr.Canvas, gr.FontCache, 1, Data, aFont, ColSpanWidth - 2 * CellPadding, 0, ref fit, ref cw);
                CellSize.Width = w;
                if (fit > 0) Data = Data.Substring(0, fit); else Data = new TRichString();
            }

            THtmlImageCache[,][] ImagesInCell;
            bool AddedTable = false;
            bool HasImages = Cache.Images.TryGetValue(r, cStart, RowSpan, ColSpan, out ImagesInCell);

            // This works, but fails in outlook 2007. Other option is not to add a table, but then the images will be aligned as the cell.
            if (HasImages)
            {
                AddedTable = true;
                string CellStyle = GetCellStyle(xls, fmt, fmt2, true, THAlign.Left, false, true, ColSpanWidth, ColSpan, true, (real)RowSpanHeight, RowSpan, true, null);
                string ImageClass = HtmlVersion == THtmlVersion.Html_32 ? String.Empty : "class ='imagecell'";
                WriteTdTag(html, FormatStr("<td {0}{1}{2}>", ImageClass, SpanStr, CellStyle));
                WriteLn(html, String.Empty);
                SpanStr = String.Empty;
                WriteImagesInCell(xls, html, r, cStart, ImagesInCell);

                WriteLn(html, "<table border='0' cellpadding='0' cellspacing='0' summary='Helper container'><tr>"); //We could get the same effect with display:table-cell, but ie ignores it. So we will add a real table here.
            }

            bool RenderAsImage = VerticalTextAsImages && (fmt.Rotation == 90 || fmt.Rotation == 180 || fmt.Rotation == 255);

            string CellComment = GetCellComment(html, xls, r, c);
            string CellCommand = GetCellCommand(html, xls, r, c);
            string MainCellStyle = GetCellStyle(xls, fmt, fmt2, !AddedTable,
                HAlign, true, AddedTable | HtmlFixes.Outlook2007CssSupport, ColSpanWidth, ColSpan, AddedTable, (real)RowSpanHeight, RowSpan, true, Data, RenderAsImage);

            string CellClassStr = HtmlVersion == THtmlVersion.Html_32 ? String.Empty : GetClass(xls, fmt, Cache);
            string CellClass = HtmlVersion == THtmlVersion.Html_32 ? String.Empty : "class='" + CellClassStr + "'";
            WriteTdTag(html, "  <td " + CellClass + MainCellStyle + SpanStr + CellComment + ">");

            //if (HasImages)
            //	WriteImagesInCell(xls, html, r, cStart, ImagesInCell);

            //we need to add a div to keep IE happy. whitespace:nobreak does not work in td inside ie if column widths are specified.
            //Also, if the row is smaller than one line of text, other browsers will make the row bigger too.

            //There are 2 options here, and no one is perfect. The first one is to use a div, in which case we will lose vertical align
            //in cells. (see http://phrogz.net/CSS/vertical-align/index.html).
            //The other one is to use a span, in which case vertical align is honored, but a row will always be as tall as necessary
            //to handle a line of text. So we cannot cut the text in a cell vertically.
            //Since I believe horizontal alignment is more important than cutting semi used rows, we will use a span. We can still use a div when the text in the row is bigger than one line.

            string NamedRangeName = GetNamedRange(r, c, Cache, xls);
            if (NamedRangeName != null) Write(html, "<span id='" + NamedRangeName + "'>"); //no risk for injection here, names can't include " or '.

            bool IsMerged = MergedRange.RowCount > 1 || MergedRange.ColCount > 1;
            bool NeedsDivInCell = CellSize.Width >= ColSpanWidth || CellSize.Height >= rh / 20f;
            if (HtmlVersion == THtmlVersion.Html_32) NeedsDivInCell = false;
            string DivId = !RowAutoHeight && CellSize.Height >= rh / 20f ? "div" : "span"; //this will not cover a multilinetext with fixed row height... but almost anything else.
            if (CanWrap && RowAutoHeight && !IsMerged) NeedsDivInCell = false;
            if (NeedsDivInCell)
            {
                Write(html, "<" + DivId + " class='" + CellClassStr + "'" +
                    GetCellAlignAndHeightTag(xls, fmt, HAlign, RowSpanHeight, Data) + CellComment + ">");
            }

            if ((fmt.Font.Style & TFlxFontStyles.Subscript) != 0) Write(html, "<sub>");
            if ((fmt.Font.Style & TFlxFontStyles.Superscript) != 0) Write(html, "<sup>");

            int HLinkPos;
            THyperLink Hlink = null;
            if (Cache.Hyperlinks.TryGetValue(FlxHash.MakeHash(r, c), out HLinkPos)) //Here we use c, not cStart.
            {
                Hlink = xls.GetHyperLink(HLinkPos);
                if (Hlink != null && Hlink.LinkType == THyperLinkType.URL)
                {
                    Write(html, FormatStr("<a href=\"{0}\" {1}>",
                        EncodeHyperlink(html, Hlink.Text), HyperlinkSuppressBorder));
                    if (HtmlVersion != THtmlVersion.Html_32)
                    {
                        Write(html, "<span class='" + CellClassStr + "' style='" + GetTextDecoration(xls, fmt.Font, null) + "'>");
                    }
                }
            }
            else if (CellCommand != "")
            {
                Write(html, FormatStr("<a href=\"{0}\" {1}>",
                        CellCommand, HyperlinkSuppressBorder));
                if (HtmlVersion != THtmlVersion.Html_32)
                {
                    Write(html, "<span class='" + CellClassStr + "' style='" + GetTextDecoration(xls, fmt.Font, null) + "'>");
                }
            }

            if (RenderAsImage)
            {
                WriteTextAsImage(xls, html, HtmlFileName, RelativeImagePath, r, c, r + MergedRange.RowCount - 1, c + MergedRange.ColCount - 1, RowHeight, ColWidth, Data, Cache, SaveImages);
            }
            else
            {
                string Before = String.Empty;
                string After = string.Empty;

                if (HtmlVersion == THtmlVersion.Html_32)
                {
                    GetFontStyle32(xls, DefaultFormat, fmt, Cache, out Before, out After);
                    Write(html, Before);
                }

                if (!aColor.Equals(ColorUtil.Empty))
                {
                    Write(html, THtmlTagCreator.StartFontColor(aColor, HtmlStyle));
                }


                Write(html, Data.ToHtml(xls, fmt, HtmlVersion, HtmlStyle, html.Encoding, Cache.OnHtmlFont));

                if (HtmlVersion == THtmlVersion.Html_32)
                {
                    Write(html, After);
                }
                if (!aColor.Equals(ColorUtil.Empty))
                {
                    Write(html, THtmlTagCreator.EndFontColor(HtmlStyle));
                }
            }

            if (Hlink != null && Hlink.LinkType == THyperLinkType.URL)
            {
                if (HtmlVersion != THtmlVersion.Html_32) Write(html, "</span></a>"); else Write(html, "</a>");
            }
            else if (CellCommand != "")
            {
                if (HtmlVersion != THtmlVersion.Html_32) Write(html, "</span></a>"); else Write(html, "</a>");
            }

            if ((fmt.Font.Style & TFlxFontStyles.Superscript) != 0) Write(html, "</sup>");
            if ((fmt.Font.Style & TFlxFontStyles.Subscript) != 0) Write(html, "</sub>");

            if (NeedsDivInCell) Write(html, FormatStr("</{0}>", DivId));
            if (NamedRangeName != null) Write(html, "</span>");

            if (AddedTable)
            {
                WriteLn(html, "</td></tr></table>");
            }
            WriteLn(html, "</td>");

            LastUsedCol = cStart + ColSpan - 1;
            if (ColSpan > 1)
            {
                cIndex = xls.ColToIndex(r, LastUsedCol + 1) - 1;  //-1 because it will be incremented after this. We need to do it this way to handle both situations, cell at c exsists and not. If it exists, it will return c. If it doesnt, it will return something bigger than c, and when substracted one we will get something smaller.
            }
        }

        private string GetNamedRange(int r, int c, TExportHtmlCache Cache, ExcelFile xls)
        {
            string NameId;
            TXlsNamedRange NamedRange;
            if (Cache.Names == null || !Cache.Names.TryGetValue(new TOneCellRef(r, c), out NamedRange)) return null;
            NameId = NamedRange.Name;
            if (Parent != null)
            {
                NamedRangeExportEventArgs e = new NamedRangeExportEventArgs(xls, xls.ActiveSheet, r, c, NamedRange, NameId);
                Parent.OnNamedRangeExport(e);
                NameId = e.NameId;
            }
            return NameId;
        }

        private TXlsCellRange CellMergedBounds(int r, int c, TExportHtmlCache Cache)
        {
            return Cache.MergedCellsInColumn.GetRange(r, c);
        }

        private void WriteTextAsImage(ExcelFile xls, TextWriter html, string FileName, string ImagePath, int r1, int c1, int r2, int c2, real RowHeight, real ColWidth, TRichString Data, TExportHtmlCache Cache, bool SaveImages)
        {
            using (Image Img = xls.RenderCells(r1, c1, r2, c2, true, Images.Props.ImageResolution, Images.Props.SmoothingMode, Images.Props.InterpolationMode, Images.Props.AntiAliased))
            {
                string ImageFile;
                string ImageLink;
                GetImageFilename(xls, "_text_" + r1.ToString() + "_" + c1.ToString(), FileName, ImagePath, 0, FHtmlFileFormat, Images.SavedImagesFormat, EngineRuns, out ImageFile, out ImageLink);

                ImageInformationEventArgs e = new ImageInformationEventArgs(xls, -1, null, null, ImageFile, ImageLink, Data.ToString(), null, Images.SavedImagesFormat);
                Parent.OnGetImageInformation(e);

                string ImageCid = null;
                string FullImageLink = e.ImageLink;
                if (e.ImageLink != null)
                {
                    if (FHtmlFileFormat == THtmlFileFormat.MHtml && UseContentId)
                    {
                        ImageCid = GetImageId();
                        FullImageLink = GetContentId(ImageCid);
                    }
                    else
                    {
                        FullImageLink = EncodeUrl(e.ImageLink);
                    }
                    ColWidth = Img.Width / Images.Props.ImageResolution * FlexCelRender.DispMul; //This is explicitly recalculated here, since the image might look bad due to rounding errors. It might not be exactly one cell size.
                    RowHeight = Img.Height / Images.Props.ImageResolution * FlexCelRender.DispMul;

                    StartImageHyperlink(html, e.HyperLink);

                    if (HtmlVersion == THtmlVersion.Html_32)
                    {
                        WriteLn(html, FormatStr("  <img src='{0}' width='{1}' height='{2}' alt='{3}' border='0' {4}",
                            FullImageLink, Img.Width, Img.Height, EncodeAsHtml(html, e.AlternateText, TEnterStyle.Char10),
                            EndOfTag));
                    }
                    else
                    {
                        WriteLn(html, FormatStr("  <img src='{0}' width='{1}' height='{2}' alt=\"{3}\" " +
                            "style='display: block; border:none; width:{4" + PointsFormat + ";height:{5" + PointsFormat + ";' {6}",
                            FullImageLink, Img.Width, Img.Height, EncodeAsHtml(html, e.AlternateText, TEnterStyle.Char10),
                            ColWidth, RowHeight,
                            EndOfTag));
                    }

                    EndImageHyperlink(html, e.HyperLink);
                }

                if (FHtmlFileFormat == THtmlFileFormat.Html)
                {
                    SaveImageEventArgs sie = new SaveImageEventArgs(xls, -1, null, e.ImageFile, e.ImageLink, e.AlternateText, e.SavedImageFormat, Img);
                    Parent.OnSaveImage(sie);

                    if (!sie.Processed)
                    {
                        if (SaveImages)
                        {
                            SaveImage(Img, e.ImageStream, e.ImageFile, e.SavedImageFormat);
                        }
                    }
                }
                else
                {
                    Cache.CellImages.Add(new THtmlCellImageCache(r1, c1, FullImageLink, IsAbsoluteUrl(e.ImageLink), e.SavedImageFormat, ImageCid));
                }
            }
        }

        string HyperlinkSuppressBorder
        {
            get
            {
                return HtmlVersion == THtmlVersion.Html_32 ? "" : "style='text-decoration: none;'";
            }
        }

        private void StartImageHyperlink(TextWriter html, string HyperLink)
        {
            if (HyperLink != null && (FHidePrintObjects & THidePrintObjects.Hyperlynks) == 0)
            {
                Write(html, FormatStr("  <a href=\"{0}\" {1}>", EncodeHyperlink(html, HyperLink), HyperlinkSuppressBorder));
            }
        }

        private void EndImageHyperlink(TextWriter html, string HyperLink)
        {
            if (HyperLink != null && (FHidePrintObjects & THidePrintObjects.Hyperlynks) == 0)
            {
                Write(html, "  </a>");
            }
        }


        private string GetCellComment(TextWriter html, ExcelFile xls, int r, int c)
        {
            if ((FHidePrintObjects & THidePrintObjects.Comments) != 0) return string.Empty;
            TRichString Comment = xls.GetComment(r, c);
            if (Comment.Value.Length == 0) return String.Empty; //Title is not strictly in 3.2, but it doesn't hurt.
            return FormatStr(" title = \"{0}\"", EncodeAsHtml(html, Comment.Value, TEnterStyle.Char10)); //no rich text here since it is not supported in "title"
        }
        private string GetCellCommand(TextWriter html, ExcelFile xls, int r, int c)
        {
            if ((FHidePrintObjects & THidePrintObjects.Comments) != 0) return string.Empty;
            TRichString Comment = xls.GetComment(r, c);
            if (Comment.Value.Length == 0) return String.Empty; //Title is not strictly in 3.2, but it doesn't hurt.
            string command = EncodeAsHtml(html, Comment.Value, TEnterStyle.Char10);
            if (command.Contains("TT_XLB_EB"))
            {
                return "tvcqd:" + command.Replace("=TT_XLB_EB", "TT_XLB_EB");
            }
            return FormatStr(" title = \"{0}\"", EncodeAsHtml(html, Comment.Value, TEnterStyle.Char10)); //no rich text here since it is not supported in "title"
        }
        private void WriteImagesInCell(ExcelFile xls, TextWriter html, int row, int col, THtmlImageCache[,][] ImagesInCell)
        {
            if (ImagesInCell != null)
            {
                real AcumHeight = 0;
                for (int r = 0; r < ImagesInCell.GetLength(0); r++)
                {
                    real AcumWidth = 0;
                    for (int c = 0; c < ImagesInCell.GetLength(1); c++)
                    {
                        if (ImagesInCell[r, c] != null)
                        {
                            foreach (THtmlImageCache Img in ImagesInCell[r, c])
                            {
                                WriteOneImage(html, AcumHeight, AcumWidth, Img);
                            }
                        }
                        AcumWidth += RealColWidth(xls, col + c);
                    }
                    AcumHeight += RealRowHeight(xls, row + r);
                }
            }
        }

        private void WriteOneImage(TextWriter html, real AcumHeight, real AcumWidth, THtmlImageCache Img)
        {

            StartImageHyperlink(html, Img.HyperLink);
            if (HtmlVersion == THtmlVersion.Html_32)
            {
                WriteHtml32Image(html, AcumHeight, AcumWidth, Img);
            }
            else
            {
                WriteStyledImage(html, AcumHeight, AcumWidth, Img);
            }

            EndImageHyperlink(html, Img.HyperLink);
        }

        private void WriteHtml32Image(TextWriter html, float AcumHeight, float AcumWidth, THtmlImageCache Img)
        {
            WriteLn(html, FormatStr("  <img src='{0}' width='{1}' height='{2}' alt='{3}' border='0' {4}",
                Img.Url, Img.SizePixels.Width, Img.SizePixels.Height, EncodeAsHtml(html, Img.AltText, TEnterStyle.Char10),
                EndOfTag));
        }

        private void WriteStyledImage(TextWriter html, real AcumHeight, real AcumWidth, THtmlImageCache Img)
        {
            //Ie has a bug here, so if we add a div, we need to add z-index to the div too.
            //see http://annevankesteren.nl/2005/06/z-index
            //We finally are not using the div, by using margin-top and margin-left instead of top and left.
            //Write(html, FormatStr("<div class = 'imagediv' style = 'z-index:{0}'>", Img.ZOrder)); 

            string ShapeStyle = String.Empty;
            if (HtmlFixes.Outlook2007CssSupport)
            {
                Guid g = Guid.NewGuid();
                string PicId = g.ToString("N", CultureInfo.InvariantCulture);
                string ShapeId = "picture" + PicId + Img.PictureId.ToString(CultureInfo.InvariantCulture); //We add the GUID here to ensure the id is different for different files.
                AddOutlookFix(html, Img.Url, EncodeAsHtml(html, Img.AltText, TEnterStyle.Char10), Img.Origin.X, Img.Origin.Y, Img.Dimensions.Width, Img.Dimensions.Height, Img.ZOrder, ShapeId);
                ShapeStyle = FormatStr("v:shapes='{0}'", ShapeId);
            }

            FixIe6Transp(html, AcumHeight, AcumWidth, Img);

            if (HtmlFixes.Outlook2007CssSupport) StartOutlookIgnore(html);

            WriteLn(html, FormatStr("  <img src='{0}' width='{1}' height='{2}' alt=\"{3}\" " +
                //"style='position: absolute; top: {4:0.##}pt; left: {5:0.##}pt; width: {6:0.##}pt; height: {7:0.##}pt; z-index:{8};' {9}",
                "class = 'imagediv' style='margin-top: {4" + PointsFormat + "; margin-left: {5" + PointsFormat + "; z-index:{8};" +
                "width:{9" + PointsFormat + ";height:{10" + PointsFormat + ";' {11} {12}",
                Img.Url, Img.SizePixels.Width, Img.SizePixels.Height, EncodeAsHtml(html, Img.AltText, TEnterStyle.Char10),
                Img.Origin.Y + AcumHeight, Img.Origin.X + AcumWidth, Img.Dimensions.Width, Img.Dimensions.Height, Img.ZOrder,
                Img.Dimensions.Width, Img.Dimensions.Height,
                ShapeStyle, EndOfTag));
            //Write(html, "</div>");

            if (HtmlFixes.Outlook2007CssSupport) EndOutlookIgnore(html);
        }

        private void FixIe6Transp(TextWriter html, real AcumHeight, real AcumWidth, THtmlImageCache Img)
        {
            bool NeedsIe6Fix = HtmlFixes.IE6TransparentPngSupport && Img.SavedImageFormat == THtmlImageFormat.Png;
            if (NeedsIe6Fix)
            {
                //fix from http://www.howtocreate.co.uk/alpha.html
                WriteLn(html, "<!--[if gte IE 5.5]>");
                WriteLn(html, "<![if lt IE 7]>");
                WriteLn(html, FormatStr("  <span " +
                    "style='margin-top: {1" + PointsFormat + "; margin-left: {2" + PointsFormat + "; z-index:{3};" +
                    "{{filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(src={0})}};width:{4" + PointsFormat + ";height:{5" + PointsFormat + ";position:absolute;" +
                    "'></span>",
                    Img.Url + ",sizingMethod=scale", Img.Origin.Y + AcumHeight, Img.Origin.X + AcumWidth, Img.ZOrder,
                    Img.Dimensions.Width, Img.Dimensions.Height
                    ));
                WriteLn(html, "<![endif]>");
                WriteLn(html, "<![endif]-->");
            }
        }

        private string FixIe6TranspBegin(real Height, real Width, string ImageUrl, THtmlImageFormat SavedImageFormat)
        {
            StringBuilder Result = new StringBuilder();

            bool NeedsIe6Fix = HtmlFixes.IE6TransparentPngSupport && SavedImageFormat == THtmlImageFormat.Png;
            if (NeedsIe6Fix)
            {
                //fix from http://www.howtocreate.co.uk/alpha.html
                Result.Append("<!--[if gte IE 5.5]>");
                Result.Append("<![if lt IE 7]>");
                Result.Append(FormatStr("<span " +
                    "style='display:inline-block;{{filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(src={0})}};width:{1" + PointsFormat + ";height:{2" + PointsFormat +
                    "'>",
                    ImageUrl, Width, Height
                    ));
                Result.Append("<![endif]>");
                Result.Append("<![endif]-->");
                Result.Append("<!--[if gte IE 7]><!-->");
            }

            return Result.ToString();
        }

        private string FixIe6TranspEnd(THtmlImageFormat SavedImageFormat)
        {
            StringBuilder Result = new StringBuilder();

            bool NeedsIe6Fix = HtmlFixes.IE6TransparentPngSupport && SavedImageFormat == THtmlImageFormat.Png;
            if (NeedsIe6Fix)
            {
                Result.Append("<!--><![endif]-->");

                Result.Append("<!--[if gte IE 5.5]>");
                Result.Append("<![if lt IE 7]>");
                Result.Append(FormatStr("</span>"));
                Result.Append("<![endif]>");
                Result.Append("<![endif]-->");
            }

            return Result.ToString();
        }

        private void AddOutlookFix(TextWriter html, string ImagePath, string ImgAlt, real left, real top, real width, real height, real zindex, string pictureid)
        {
            WriteLn(html, FormatStr(
                "  <!--[if gte vml 1]><v:shape id=\"{6}\" type=\"#_x0000_t75\" alt=\"{0}\" style='position:absolute;" +  //type = _x0000_t75 is a square. 71 f.i. is a star.
                "margin-top: {1" + PointsFormat + "; margin-left: {2" + PointsFormat + ";" +
                "width: {3" + PointsFormat + "; height: {4" + PointsFormat + "; z-index:{5};mso-wrap-style:square;" +
                "mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text'>",
                ImgAlt, top, left, width, height, zindex, pictureid));

            WriteLn(html, FormatStr("<v:imagedata src=\"{0}\" />", ImagePath));
            WriteLn(html, FormatStr("</v:shape><![endif]-->"));
        }

        private void StartOutlookIgnore(TextWriter html)
        {
            //WriteLn(html, "<!--[if lt vml 1]>");
        }

        private void EndOutlookIgnore(TextWriter html)
        {
            //WriteLn(html, "<![endif]-->");
        }

        private void GetColClass(ExcelFile xls, int r, int c, TExportHtmlCache Cache, out string CellStyle, out string CellClass)
        {
            TFlxFormat fmt = Cache.Formats.GetCellVisibleFormatDef(xls, r, c, false);
            //sadly we cannot ommit the class even if the cell is empty, outlook 2007/word 2007 will not inherit the css. so it needs to be in the cell. 
            {
                CellClass = GetClass(xls, fmt, Cache);
                CellStyle = GetCellStyle(xls, fmt, fmt, true, THAlign.Left, false, false, 0, 1, false, 0, 1, false, null);
            }
        }

        private void WriteEmptyTdTag(TextWriter html, string s, string namedRangeName, ref int colspan, bool avoidMerge)
        {
            //We are not enabling this at the moment. It can cause borders to span because of the border collapse property of the table.
            //so we will write a <td> for each cell, even if a little more verbose.
            /*if (!avoidMerge && (last == null || last == s))
            {
                colspan++;
            }
            else*/
            {
                WriteFullTdTag(html, s, colspan, namedRangeName);
                colspan = 1;
            }

            //last = s;
        }

        private void WriteFullTdTag(TextWriter html, string s, int colspan, string namedRangeName)
        {
            if (s == null) return;

            WriteTdTag(html, s);
            if (colspan > 1) Write(html, " colspan = '" + colspan.ToString(CultureInfo.InvariantCulture) + "'");
            Write(html, ">");
            if (HtmlVersion == THtmlVersion.Html_32) Write(html, "&nbsp;");

            if (namedRangeName != null) Write(html, FormatStr("<span id='{0}'></span>", namedRangeName)); //no risk for injection here, names can't include " or '.
            WriteLn(html, "</td>");
        }

        private void FillEmptyCells(ExcelFile xls, TextWriter html, int r, int StartCol, int EndCol, TExportHtmlCache Cache)
        {
            if (EndCol - StartCol <= 1) return;

            bool PrintGridLines = xls.PrintGridLines;

            Write(html, "  ");
            int colspan = 0;

            for (int c = StartCol + 1; c < EndCol; c++)
            {
                if (xls.GetColHidden(c)) continue;
                real w1 = RealColWidth(xls, c);
                if (w1 <= 0) continue;  //Not exported.

                THtmlImageCache[,][] Images;

                string CellClass = string.Empty;
                string CellStyle = string.Empty;
                string FullCellStyle;

                if (HtmlVersion == THtmlVersion.Html_32)
                {
                    TFlxFormat fmt = Cache.Formats.GetCellVisibleFormatDef(xls, r, c, false);
                    FullCellStyle = " " + GetCellStyle(xls, fmt, fmt, true, THAlign.Left, false, false, 0, 1, false, 0, 1, false, null);
                }
                else
                {
                    GetColClass(xls, r, c, Cache, out CellStyle, out CellClass);
                    FullCellStyle = CellStyle.Length == 0 ? String.Empty : " style = '" + CellStyle + "'"; //FormatStr is a little slow, and fillemptycells is a key method.
                }
                string FullCellClass = CellClass.Length == 0 ? String.Empty : " class = '" + CellClass + "'";

                TXlsCellRange MergedRange = null;
                if (HtmlVersion == THtmlVersion.Html_32) //in html32, a merged cell might be "virtual" and not have a real cell below. This would fool this method if we don't check it. If not html32, there is no need, so we don't do it to keep this as fast as possible.
                {
                    MergedRange = CellMergedBounds(r, c, Cache);
                    if (c != MergedRange.Left || r != MergedRange.Top) continue;
                }

                string NamedRangeName = GetNamedRange(r, c, Cache, xls);

                if (Cache.Images.TryGetValue(r, c, 1, 1, out Images))
                {
                    colspan = 0;

                    if (HtmlVersion == THtmlVersion.Html_32)
                    {
                        string SpanStr = string.Empty;
                        if (!MergedRange.IsOneCell) SpanStr = FormatStr(" colspan = '{0}' rowspan ='{1}'", MergedRange.ColCount, MergedRange.RowCount);
                        WriteTdTag(html, "<td" + FullCellStyle + SpanStr + GetCellComment(html, xls, r, c) + ">");
                    }
                    else
                    {
                        if (CellClass.Length == 0)
                        {
                            WriteTdTag(html, "<td class = 'imagecell'" + GetCellComment(html, xls, r, c) + FullCellStyle + ">");
                        }
                        else
                        {
                            WriteTdTag(html, "<td" + FullCellClass + GetCellComment(html, xls, r, c) + " style = 'vertical-align:top;text-align:left;padding:0;" + CellStyle + "'>");  //We will inline the class imagecell here, since word does not like 2 classes in the same cell.
                        }
                    }

                    if (NamedRangeName != null) Write(html, FormatStr("<span id='{0}'>", NamedRangeName)); //no risk for injection here, names can't include " or '.
                    WriteImagesInCell(xls, html, r, c, Images);
                    if (NamedRangeName != null) Write(html, "</span>");
                    WriteLn(html, "</td>");
                }
                else
                {
                    WriteEmptyTdTag(html, "<td" + FullCellClass + FullCellStyle + GetCellComment(html, xls, r, c), NamedRangeName,
                        ref colspan, PrintGridLines);
                }
            }
            WriteLn(html, "");
        }

        internal void WriteEndDoc(TextWriter html)
        {
            WriteLn(html, "</html>");
        }
        #endregion

        #region Write Images
        internal static string GetImageContentType(THtmlImageFormat imageformat)
        {
            switch (imageformat)
            {
                case THtmlImageFormat.Gif: return "image/gif";
                case THtmlImageFormat.Jpeg: return "image/jpeg";
            }
            return "image/png";
        }

        private static string GetExtension(THtmlImageFormat imageformat)
        {
            switch (imageformat)
            {
                case THtmlImageFormat.Gif: return ".gif";
                case THtmlImageFormat.Jpeg: return ".jpg";
            }
            return ".png";
        }

        internal static ImageFormat GetImageFormat(THtmlImageFormat imageformat)
        {
            switch (imageformat)
            {
                case THtmlImageFormat.Gif: return ImageFormat.Gif;
                case THtmlImageFormat.Jpeg: return ImageFormat.Jpeg;
            }
            return ImageFormat.Png;
        }

        private static string GetAlternateText(int i, TShapeProperties ShProp)
        {
            TShapeOptionList ShapeOptions = ShProp.NestedOptions;
            if (ShapeOptions != null)
            {

                string AltText = ShapeOptions.AsUnicodeString(TShapeOption.wzDescription, null);
                if (AltText != null) return AltText;

                AltText = ShapeOptions.AsUnicodeString(TShapeOption.gtextUNICODE, null);
                if (AltText != null) return AltText;

                AltText = ShapeOptions.AsUnicodeString(TShapeOption.gtextRTF, ShProp.ShapeName);
            }
            return "Image " + i.ToString(CultureInfo.InvariantCulture);

        }

        private static string GetImageHyperlink(int i, TShapeProperties ShProp)
        {
            TShapeOptionList ShapeOptions = ShProp.NestedOptions;
            if (ShapeOptions != null)
            {
                THyperLink hl = ShapeOptions.AsHyperLink(TShapeOption.pihlShape, null);
                if (hl == null || hl.Text == null || hl.Text.Length == 0) return null;
                return hl.Text;
            }
            return null;
        }

        private void GetImageFilename(ExcelFile xls, string ImagePrefix, string FileName, string ImagePath, int i, THtmlFileFormat FileFormat, THtmlImageFormat SavedImagesFormat, int aEngineRuns, out string ImageFile, out string ImageLink)
        {
            string ImageFileName = string.Empty;

            switch (Images.Props.ImageNaming)
            {
                case TImageNaming.Guid:
                    ImageFileName = Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture) + GetExtension(SavedImagesFormat);
                    break;

                default:
                    string RunId = aEngineRuns == 0 ? String.Empty : String.Format(CultureInfo.InvariantCulture, "_{0}", aEngineRuns); //So images in different places will have different names.
                    ImageFileName = RunId + ImagePrefix + i.ToString(CultureInfo.InvariantCulture) + GetExtension(SavedImagesFormat);
                    ImageFileName = Path.GetFileNameWithoutExtension(FileName) + ImageFileName;
                    break;
            }

            GetImageLinks(FileName, ImagePath, ImageFileName, out ImageFile, out ImageLink);
        }

        private static void GetImageLinks(string FileName, string ImagePath, string ImageFileName, out string ImageFile, out string ImageLink)
        {
            ImageFile = null;
            ImageLink = null;
            if (FileName != null)
            {
                if (ImagePath != null)
                {
                    ImageLink = Path.Combine(ImagePath, ImageFileName);
                }
                else
                {
                    ImageLink = ImageFileName;
                }

                ImageFile = Path.Combine(Path.GetDirectoryName(FileName), ImageLink);
                ImageLink = EncodeUrl(ImageLink);
            }
        }

        private static string GetContentId(string cid)
        {
            return "cid:" + cid;
        }

        private string GetImageId()
        {
            return Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture);
        }


        private void CacheMergedCells(ExcelFile xls, TExportHtmlCache Cache)
        {
            Cache.MergedCellsInColumn = new TMergedCellsInColumn();
            for (int i = 1; i <= xls.CellMergedListCount; i++)
            {
                TXlsCellRange Range = xls.CellMergedList(i);
                Cache.MergedCellsInColumn.Add(Range.Top, Range.Left, Range.Bottom, Range.Right);
            }

            if (HtmlVersion == THtmlVersion.Html_32 && Cache.Images != null)
            {
                MergeCellsUnderImages(xls, Cache);
            }
        }

        private void MergeCellsUnderImages(ExcelFile xls, TExportHtmlCache Cache)
        {
            //Merge Cells below images, so they look nicer. We don't have floating images here.
            foreach (THtmlImageCache[] Imgs in Cache.Images.Values)
            {
                foreach (THtmlImageCache Img in Imgs)
                {
                    Cache.MergedCellsInColumn.Add(Img.r1, Img.c1, Img.r2, Img.c2);
                }
            }
        }

        public void ExtractImages(ExcelFile xls, TExportHtmlCache Cache, string FileName, string ImagePath, bool SaveImages)
        {
            if ((FHidePrintObjects & THidePrintObjects.Images) != 0) return;
            bool MergeCells = HtmlVersion == THtmlVersion.Html_32 && MergeImagesInHTML32;

            GetIndividualImages(xls, Cache, FileName, ImagePath, SaveImages & !MergeCells);
            CacheMergedCells(xls, Cache);

            if (HtmlVersion == THtmlVersion.Html_32)
            {
                if (MergeImagesInHTML32) RemergeImages(xls, Cache); else RegroupImages(Cache);
            }

        }

        private void RemergeImages(ExcelFile xls, TExportHtmlCache Cache)
        {
            /*PENDING: This method should discard old individual images, and call RenderCells to have images that fill the 
             * full merged ranges. It should also call a new event to allow naming of the new images. As it is now, MergeImagesInHTML32 
             * is always false, so this method is never called.
             * 
            THtmlImageCacheList NewImages = new THtmlImageCacheList();
            foreach (long rowcol in Cache.Images.Keys)
            {
                int row; int col;
                FlxConsts.UnHash(rowcol, out row, out col);
                TXlsCellRange range = Cache.MergedCellsInColumn.GetRange(row, col);

                using (Image im = xls.RenderCells(range.Top, range.Left, range.Bottom, range.Right, true))
                {
                    NewImages.Add(range.Top, range.Left, img);
                }
            }
            Cache.Images = NewImages;*/
        }

        private static void RegroupImages(TExportHtmlCache Cache)
        {
            THtmlImageCacheList NewImages = new THtmlImageCacheList();
            foreach (long rowcol in Cache.Images.Keys)
            {
                int row; int col;
                FlxHash.UnHash(rowcol, out row, out col);
                TXlsCellRange range = Cache.MergedCellsInColumn.GetRange(row, col);

                foreach (THtmlImageCache img in (THtmlImageCache[])Cache.Images[rowcol])
                {
                    NewImages.Add(range.Top, range.Left, img);
                }
            }

            Cache.Images = NewImages;
        }

        private void GetIndividualImages(ExcelFile xls, TExportHtmlCache Cache, string FileName, string ImagePath, bool SaveImages)
        {
            int ObjCount;
            bool IsChart = xls.SheetType == TSheetType.Chart;
            if (IsChart)
            {
                ObjCount = 1;
            }
            else
            {
                ObjCount = xls.ObjectCount;
            }

            for (int i = 1; i <= ObjCount; i++)
            {
                TShapeProperties ShProp;
                if (IsChart)
                {
                    ShProp = new TShapeProperties();
                    ShProp.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 1, 0, 1, 0, 1, 0, 1, 0);
                }
                else
                {
                    ShProp = xls.GetObjectProperties(i, true);
                }
                if (!ShProp.Print || !ShProp.Visible || ShProp.ObjectType == TObjectType.Comment) continue;

                RectangleF Dimensions;
                PointF Origin;
                Size ImageSizePixels;
                using (Image Img = GetImage(xls, SaveImages || Parent.HasSaveImageEvent, i, ShProp, out Dimensions, out Origin, out ImageSizePixels))
                {
                    if (Dimensions.Width > 0 && Dimensions.Height > 0)
                    {
                        string ImageFile;
                        string ImageLink;
                        TClientAnchor Anchor = ShProp.NestedAnchor;
                        GetImageFilename(xls, "_image", FileName, ImagePath, i, FHtmlFileFormat, Images.SavedImagesFormat, EngineRuns, out ImageFile, out ImageLink);
                        ImageInformationEventArgs e = new ImageInformationEventArgs(xls, i, ShProp, null, ImageFile, ImageLink, GetAlternateText(i, ShProp), GetImageHyperlink(i, ShProp), Images.SavedImagesFormat);
                        Parent.OnGetImageInformation(e);

                        if (e.ImageLink != null)
                        {
                            string ImageCid = null;
                            string FullImageLink = e.ImageLink;
                            if (FHtmlFileFormat == THtmlFileFormat.MHtml && UseContentId)
                            {
                                ImageCid = GetImageId();
                                FullImageLink = GetContentId(ImageCid);
                            }
                            else
                            {
                                FullImageLink = EncodeUrl(e.ImageLink);
                            }
                            Cache.Images.Add(Anchor.Row1, Anchor.Col1, new THtmlImageCache(Anchor.Row1, Anchor.Col1, Anchor.Row2, Anchor.Col2,
                                FullImageLink, IsAbsoluteUrl(e.ImageLink), i,
                                new Size((int)Math.Round(ImageSizePixels.Width * 96f / Images.Props.ImageResolution), (int)Math.Round(ImageSizePixels.Height * 96f / Images.Props.ImageResolution)),
                                Origin, Dimensions, e.AlternateText, e.HyperLink, i, e.SavedImageFormat, ImageCid));
                        }

                        if (Img != null)
                        {
                            SaveImageEventArgs sie = new SaveImageEventArgs(xls, i, ShProp, e.ImageFile, e.ImageLink, e.AlternateText, e.SavedImageFormat, Img);
                            Parent.OnSaveImage(sie);

                            if (!sie.Processed && SaveImages)
                            {
                                SaveImage(Img, e.ImageStream, e.ImageFile, e.SavedImageFormat);
                            }
                        }
                    }
                }
            }
        }

        private Image GetImage(ExcelFile xls, bool SaveImages, int PictureId, TShapeProperties ShProp, out RectangleF Dimensions, out PointF Origin, out Size ImageSizePixels)
        {
            if (xls.SheetType == TSheetType.Chart)
            {
                using (FlexCelImgExport ChartImage = new FlexCelImgExport(xls, false))
                {
                    ChartImage.HidePrintObjects = FHidePrintObjects;
                    ChartImage.Resolution = Images.Props.ImageResolution;

                    Origin = new PointF(0, 0);
                    TPaperDimensions ps = ChartImage.GetRealPageSize();
                    Dimensions = new RectangleF(0, 0, ps.Width / 96f * FlexCelRender.DispMul, ps.Height / 96f * FlexCelRender.DispMul);
                    int wPix = (int)Math.Ceiling(ps.Width * Images.Props.ImageResolution / 96f) + 1;
                    int hPix = (int)Math.Ceiling(ps.Height * Images.Props.ImageResolution / 96f) + 1;
                    ImageSizePixels = new Size(wPix, hPix);


                    if (!SaveImages) return null;

                    Color ImgBkg = Images.Props.ImageBackground == ColorUtil.Empty ? Color.White : Images.Props.ImageBackground; //We don't want transparency here, this is a sheet.
                    using (TBitmapCreator bmp = new TBitmapCreator(Images.Props.ImageResolution, Images.Props.SmoothingMode, Images.Props.AntiAliased, Images.Props.InterpolationMode, ImgBkg, wPix, hPix))
                    {
                        TImgExportInfo ExportInfo = null;
                        ChartImage.ExportNext(bmp.ImgGraphics, ref ExportInfo);
                        return bmp.Img;
                    }
                }
            }

            return xls.RenderObject(PictureId, Images.Props.ImageResolution, ShProp, Images.Props.SmoothingMode, Images.Props.InterpolationMode, Images.Props.AntiAliased,
                                           SaveImages, Images.Props.ImageBackground, out Origin, out Dimensions, out ImageSizePixels);
        }

        public void SaveMHTMLImages(ExcelFile xls, TextWriter HtmlStream, TExportHtmlCache Cache, TMimeWriter MHTML)
        {
            ExportDrawingObjects(xls, HtmlStream, Cache, MHTML);
            ExportCellImages(xls, HtmlStream, Cache, MHTML);
            ExportHeaderImages(xls, HtmlStream, Cache, MHTML);
        }

        private void ExportDrawingObjects(ExcelFile xls, TextWriter HtmlStream, TExportHtmlCache Cache, TMimeWriter MHTML)
        {
            foreach (THtmlImageCache[] images in Cache.Images.Values)
            {
                foreach (THtmlImageCache image in images)
                {

                    TShapeProperties ShProp = xls.SheetType == TSheetType.Chart ? new TShapeProperties() : xls.GetObjectProperties(image.PictureId, true);
                    if (!ShProp.Print || !ShProp.Visible || ShProp.ObjectType == TObjectType.Comment) continue;

                    RectangleF Dimensions;
                    PointF Origin;
                    Size ImageSizePixels;
                    if (image.UrlIsAbsolute) continue;  //might be a link to an external file (http://...) This works in ie, not in opera.
                    using (Image Img = GetImage(xls, true, image.PictureId, ShProp, out Dimensions, out Origin, out ImageSizePixels))
                    {
                        Uri FileUrl = image.UniqueId == null ? GetFileUrl(image.Url) : null;
                        MHTML.AddPartHeader(HtmlStream, GetImageContentType(image.SavedImageFormat), TContentTransferEncoding.Base64, FileUrl, image.UniqueId, null);
                        using (MemoryStream ms = new MemoryStream())
                        {
                            //Do not add to generated files here.
                            using (ImageContainer Img2 = PrepareImage(Img, image.SavedImageFormat))
                            {
                                Img2.Img.Save(ms, GetImageFormat(image.SavedImageFormat));
                            }
                            MHTML.WriteBase64(HtmlStream, ms.ToArray());
                        }
                        MHTML.EndPart(HtmlStream);
                    }
                }
            }
        }

        private void ExportCellImages(ExcelFile xls, TextWriter HtmlStream, TExportHtmlCache Cache, TMimeWriter MHTML)
        {
            foreach (THtmlCellImageCache image in Cache.CellImages)
            {
                if (image.UrlIsAbsolute) continue;  //might be a link to an external file (http://...) This works in ie, not in opera.
                using (Image Img = xls.RenderCells(image.Row, image.Col, image.Row, image.Col, true, Images.Props.ImageResolution, Images.Props.SmoothingMode, Images.Props.InterpolationMode, Images.Props.AntiAliased))
                {
                    Uri FileUrl = image.UniqueId == null ? GetFileUrl(image.Url) : null;
                    MHTML.AddPartHeader(HtmlStream, GetImageContentType(image.SavedImageFormat), TContentTransferEncoding.Base64, FileUrl, image.UniqueId, null);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        //Do not add to generated files here.
                        using (ImageContainer Img2 = PrepareImage(Img, image.SavedImageFormat))
                        {
                            Img2.Img.Save(ms, GetImageFormat(image.SavedImageFormat));
                        }
                        MHTML.WriteBase64(HtmlStream, ms.ToArray());
                    }
                    MHTML.EndPart(HtmlStream);
                }
            }
        }

        private void ExportHeaderImages(ExcelFile xls, TextWriter HtmlStream, TExportHtmlCache Cache, TMimeWriter MHTML)
        {
            foreach (THtmlHeaderImageCache image in Cache.HeaderImages)
            {
                if (image.UrlIsAbsolute) continue;  //might be a link to an external file (http://...) This works in ie, not in opera.
                using (Image Img = GetHeaderImg(xls, THeaderAndFooterKind.Default, image.Section))
                {
                    Uri FileUrl = image.UniqueId == null ? GetFileUrl(image.Url) : null;
                    MHTML.AddPartHeader(HtmlStream, GetImageContentType(image.SavedImageFormat), TContentTransferEncoding.Base64, FileUrl, image.UniqueId, null);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        //Do not add to generated files here.
                        using (ImageContainer Img2 = PrepareImage(Img, image.SavedImageFormat))
                        {
                            Img2.Img.Save(ms, GetImageFormat(image.SavedImageFormat));
                        }
                        MHTML.WriteBase64(HtmlStream, ms.ToArray());
                    }
                    MHTML.EndPart(HtmlStream);
                }
            }
        }

        private static string UriToString(Uri u)
        {
            return u.ToString();
        }

        internal static Uri GetFileUrl(string s)
        {
            //see: http://blogs.msdn.com/ie/archive/2006/12/06/file-uris-in-windows.aspx  (File URIs in Windows)
            //However, Opera needs a host, so we cannot use file:///, we need to use file://localhost/
            //Update: If we use: file://localhost/ outlook 2007 refuses to open the image. So this is reverted to file:///
            //Opera is not as important here, and they are more likely to fix it in a newer version anyway.
            return new Uri("file:///C:/" + s);
        }

        internal static string EncodeUrl(string aUrl)
        {
#if (FRAMEWORK20_SKIP)  //Commented out because it will not work in Mono NET 2.0, and does not add real functionality or advantages.
            aUrl = aUrl.Replace("\\", "/");
			if (Uri.IsWellFormedUriString(aUrl, UriKind.RelativeOrAbsolute)) return aUrl;
            return Uri.EscapeUriString(aUrl);
#else
            //Sadly Uri in .NET 1.1 does not support relative paths, which renders it basically useless to store the Uri. So we need to use a string, but we can use Uri here as helper.
            if (IsAbsoluteUrl(aUrl)) return new Uri(aUrl).AbsoluteUri;
            Uri BaseUri = new Uri("file:///C:/");
            Uri Result = new Uri(BaseUri, aUrl);

            //Note that ie needs the uri *not* to be an uri, by allowing chars > 255.
            //(http://blogs.msdn.com/ie/archive/2006/12/06/file-uris-in-windows.aspx  the part on non ascii characters)
            //But we are using an uri here, and .NET uri is correct (with escaped unicode characters). 
            //So there is no solution for this, unicode uris will not show in ie. If we fixed it for ie, we would be breaking it for other browsers.
            //See also http://www.w3.org/TR/html401/appendix/notes.html#h-B.2.1

#if (FRAMEWORK20)
            return BaseUri.MakeRelativeUri(Result).ToString();
#else
			return BaseUri.MakeRelative(Result);
#endif

#endif
        }

        private static bool IsAbsoluteUrl(string s)
        {
            return s.IndexOf(Uri.SchemeDelimiter) > 0;
        }

        private static ImageContainer PrepareImage(Image Img, THtmlImageFormat ImgFormat)
        {
#if (!FULLYMANAGED)
            if (ImgFormat != THtmlImageFormat.Gif) return new ImageContainer(Img, false);

            if (!FlxUtils.HasUnamanagedPermissions())
            {
                return new ImageContainer(Img, false); //Leave the gif to be (badly) converted by GDI+
            }
            return GetOptimizedImageContainer(Img);
#else
            return new ImageContainer(Img, false); //Leave the gif to be (badly) converted by GDI+
#endif
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
#if (FRAMEWORK40)
        [SecuritySafeCritical] //This will only be called by trusted people.
#endif
        private static ImageContainer GetOptimizedImageContainer(Image Img)
        {
            return new ImageContainer(OctreeQuantizer.ConvertTo256Colors(Img), true);
        }

        private void SaveImage(Image Img, Stream ImageStream, string ImageFile, THtmlImageFormat ImgFormat)
        {
            if (ImageStream != null)
            {
                using (ImageContainer Img2 = PrepareImage(Img, ImgFormat))
                {
                    Img2.Img.Save(ImageStream, GetImageFormat(ImgFormat));
                }
            }
            else if (ImageFile != null)
            {
                Parent.GeneratedFiles.AddImage(ImageFile);
                Directory.CreateDirectory(Path.GetDirectoryName(ImageFile));
                using (ImageContainer Img2 = PrepareImage(Img, ImgFormat))
                {
                    SaveImageAndRecover(ImageFile, ImgFormat, Img2);
                }
            }
        }

        private void SaveImageAndRecover(string ImageFile, THtmlImageFormat ImgFormat, ImageContainer Img2)
        {
            TSaveImageSV ImgSaver = new TSaveImageSV(Img2.Img, ImgFormat);
            ImgSaver.Save(Parent.IgnoreSharingViolations, Parent.AllowOverwritingFiles, ImageFile);
        }


        #endregion

        #region Write CSS
        private void WriteCssLink(TextWriter html, string externalCssURL)
        {
            WriteLn(html, FormatStr("<link href='{0}' title='normal' rel='stylesheet' type='text/css'" + EndOfTag, EncodeUrl(externalCssURL)));

        }

        internal void WriteCss(TextWriter html, TCssInformation CssInfo, TSheetSelector SheetSelector)
        {
            if (HtmlVersion == THtmlVersion.Html_32) return; //no CSS in 3.2

            if (CssInfo != null && CssInfo.Url != null)
            {
                WriteCssLink(html, EncodeUrl(CssInfo.Url));
                return;
            }

            WriteLn(html, "<style type = 'text/css'>");
            if (HtmlVersion == THtmlVersion.Html_401) WriteLn(html, "<!--");  //Proper XHTML CSS should not be inside comments, or the browser might decide to treat the block as a real comment, not as CSS.


            WriteGenericCssClasses(html, SheetSelector);

            foreach (string fmt in CssInfo.UsedFormats.Values)
            {
                WriteOneCss(html, fmt, CssInfo.UsedFormats.GetId(fmt));
            }

            if (HtmlVersion == THtmlVersion.Html_401) WriteLn(html, "-->");
            WriteLn(html, "</style>");

        }

        private void WriteGenericCssClasses(TextWriter css, TSheetSelector SheetSelector)
        {
            if (HtmlVersion == THtmlVersion.Html_32) return;

            string Padding = FormatStr("{0:0.##}", CellPadding);
            WriteLn(css, "table." + MainTableClass + " td {overflow:hidden;padding: 0 " + Padding + "pt}"); // Adding a padding in the cell will make ie calculate the wrong dimensions when using table-layout:fixed. we will fix it by making the first row empty with padding 0.
            WriteLn(css, " .imagediv {position:absolute;border:none}");
            WriteLn(css, " table td.imagecell {vertical-align:top;text-align:left;padding:0}"); //We add a "table and "td" before class def to give this rule priority over the first rule in this block.
            WriteLn(css, " ." + HeadingClass + " {" + FHeadingStyle + "}");


            if (SheetSelector != null)
            {
                try
                {
                    SheetSelector.AddInternals(this, css);
                    SheetSelector.WriteCssClasses();
                }
                finally
                {
                    SheetSelector.RestoreInternals();
                }

            }
        }

        private void WriteOneCss(TextWriter css, string fmtStr, int fmtId)
        {
            WriteLn(css, " ." + ClassPrefix + fmtId + " {");
            WriteLn(css, fmtStr);
        }

        internal static void SearchUsedXF(ExcelFile xls, TXlsCellRange[] PagePrintRange, TExportHtmlCache Cache)
        {
            foreach (TXlsCellRange rng in PagePrintRange)
            {
                SearchUsedXF(xls, rng, Cache);
            }
        }

        internal static void SearchUsedXF(ExcelFile xls, TXlsCellRange PagePrintRange, TExportHtmlCache Cache)
        {
            if (xls == null) return;
            if (Cache.CssInfo != null && Cache.CssInfo.Url != null) return; //CSS is external, no need to cache.


            for (int c = PagePrintRange.Left; c <= PagePrintRange.Right; c++)
            {
                int xf = xls.GetColFormat(c);
                if (xf <= 0) xf = xls.DefaultFormatId;

                Cache.UsedFormats.Add(xls, Cache.OnHtmlFont, xls.GetFormat(xf));

            }

            for (int r = PagePrintRange.Top; r <= PagePrintRange.Bottom; r++)
            {
                int xf = xls.GetRowFormat(r);
                if (xf >= 0)
                {
                    Cache.UsedFormats.Add(xls, Cache.OnHtmlFont, xls.GetFormat(xf));
                }

                int Rows = xls.ColCountInRow(r);
                for (int cIndex = 1; cIndex <= Rows; cIndex++)
                {
                    int c = xls.ColFromIndex(r, cIndex);
                    if (c < PagePrintRange.Left) continue;
                    if (c > PagePrintRange.Right) break;

                    TFlxFormat fmt = Cache.Formats.GetCellVisibleFormatDef(xls, r, c, false);
                    Cache.UsedFormats.Add(xls, Cache.OnHtmlFont, fmt);
                }
            }
        }

        #endregion

        #region Write Headers and Footers
        private void WriteHeaders(ExcelFile xls, TextWriter html, string FileName, string ImagePath, TExportHtmlCache Cache, bool SaveImages, real TableWidth)
        {
            WriteHeaderOrFooter(xls, html, FileName, ImagePath, xls.PageHeader, Cache, true, SaveImages, TableWidth);
        }

        private void WriteFooters(ExcelFile xls, TextWriter html, string FileName, string ImagePath, TExportHtmlCache Cache, bool SaveImages, real TableWidth)
        {
            WriteHeaderOrFooter(xls, html, FileName, ImagePath, xls.PageFooter, Cache, false, SaveImages, TableWidth);
        }

        private void WriteHeaderOrFooter(ExcelFile xls, TextWriter html, string FileName, string ImagePath, string text, TExportHtmlCache Cache,
            bool Header, bool SaveImages, real TableWidth)
        {
            if (HtmlVersion == THtmlVersion.Html_32) return;

            if (text == null || text.Trim().Length == 0) return;

            TXlsMargins m = xls.GetPrintMargins();
            double h = Header ? m.Top : m.Bottom;

            WriteLn(html, FormatStr("<div style = 'width:{0" + PointsFormat, TableWidth) + ";position:relative;" +
                FormatStr("height:{0" + PointsFormat + ";", h * 72f) +
                "'>");

            string Left = String.Empty;
            string Center = String.Empty;
            string Right = String.Empty;
            xls.FillPageHeaderOrFooter(text, ref Left, ref Center, ref Right);

            int hpos = Header ? 0 : 3;
            WriteHeaderSection(xls, html, FileName, ImagePath, "left", Left, Cache, Header, "l", SaveImages, THeaderAndFooterKind.Default, ((THeaderAndFooterPos)hpos));
            WriteHeaderSection(xls, html, FileName, ImagePath, "center", Center, Cache, Header, "c", SaveImages, THeaderAndFooterKind.Default, ((THeaderAndFooterPos)hpos + 1));
            WriteHeaderSection(xls, html, FileName, ImagePath, "right", Right, Cache, Header, "r", SaveImages, THeaderAndFooterKind.Default, ((THeaderAndFooterPos)hpos + 2));

            WriteLn(html, "</div>");

        }

        private Image GetHeaderImg(ExcelFile xls, THeaderAndFooterKind Kind, THeaderAndFooterPos Section)
        {
            THeaderOrFooterImageProperties ImgProp = xls.GetHeaderOrFooterImageProperties(Kind, Section);
            if (ImgProp == null || ImgProp.Anchor == null) return null;
            real imgw = ImgProp.Anchor.Width / 96F * FlexCelRender.DispMul;
            real imgh = ImgProp.Anchor.Height / 96F * FlexCelRender.DispMul;
            RectangleF Coords = new RectangleF(0, 0, imgw, imgh);

            using (MemoryStream ImgData = new MemoryStream())
            {
                TXlsImgType ImageType = TXlsImgType.Unknown;
                xls.GetHeaderOrFooterImage(Kind, Section, ref ImageType, ImgData);
                if (ImgData != null)
                {
                    int wPix = (int)Math.Ceiling(imgw * Images.Props.ImageResolution / FlexCelRender.DispMul) + 1;
                    int hPix = (int)Math.Ceiling(imgh * Images.Props.ImageResolution / FlexCelRender.DispMul) + 1;

                    using (TBitmapCreator bmp = new TBitmapCreator(Images.Props.ImageResolution, Images.Props.SmoothingMode, Images.Props.AntiAliased, Images.Props.InterpolationMode, Images.Props.ImageBackground, wPix, hPix))
                    {
                        GdiPlusGraphics Canvas = new GdiPlusGraphics(bmp.ImgGraphics);
                        //No need to create sformat, since we are not going to write text here.					
                        DrawShape.DrawOneImage(Canvas, ImgProp.CropArea, ImgProp.TransparentColor, ImgProp.Brightness, ImgProp.Contrast,
                            ImgProp.Gamma, ColorUtil.Empty, ImgProp.BiLevel, ImgProp.Grayscale, Coords, ImgData);

                        return bmp.Img;
                    }
                }
            }

            return null;
        }

        private void WriteHeaderSection(ExcelFile xls, TextWriter html, string FileName, string ImagePath,
            string align, string Left, TExportHtmlCache Cache, bool Header, string section, bool SaveImages, THeaderAndFooterKind Kind, THeaderAndFooterPos sectionloc)
        {
            if (Left == null || Left.Trim().Length == 0) return;

            TXlsMargins m = xls.GetPrintMargins();

            string vpos = String.Empty;
            if (Header)
            {
                vpos = FormatStr("top:{0" + PointsFormat + ";", m.Header * 72f);
            }
            else
            {
                vpos = FormatStr("bottom:{0" + PointsFormat + ";", m.Footer * 72f);
            }

            string hdr = Header ? "header" : "footer";

            HtmlFontEventArgs ex = new HtmlFontEventArgs(xls, new TFlxFont(), "Arial,sans-serif");
            if (Cache.OnHtmlFont != null) Cache.OnHtmlFont.DoHtmlFont(ex);

            real lmargin = 0.75F - (real)m.Left;
            if (lmargin < 0 || (sectionloc != THeaderAndFooterPos.HeaderLeft && sectionloc != THeaderAndFooterPos.FooterLeft)) lmargin = 0;

            WriteLn(html, "<div style = 'font-family:" + ex.FontFamily + ";font-size:10pt;width:100%;position:absolute;" +
                vpos + "left:0;text-align:" +
                align +
                ";padding: 0pt " + FormatStr("{0" + PointsFormat + ";", lmargin * FlxConsts.DispMul) +
                "'>");

            string ImageTag = null;
            if ((FHidePrintObjects & THidePrintObjects.Images) == 0)
            {
                using (Image Img = GetHeaderImg(xls, Kind, sectionloc))
                {
                    if (Img != null)
                    {
                        string ImageFile;
                        string ImageLink;
                        GetImageFilename(xls, "_" + hdr + "_" + section, FileName, ImagePath, 0, FHtmlFileFormat, Images.SavedImagesFormat, EngineRuns, out ImageFile, out ImageLink);

                        ImageInformationEventArgs e = new ImageInformationEventArgs(xls, -1, null, null, ImageFile, ImageLink, hdr, null, Images.SavedImagesFormat);
                        Parent.OnGetImageInformation(e);

                        string ImageCid = null;
                        string FullImageLink = e.ImageLink;
                        if (e.ImageLink != null)
                        {
                            if (FHtmlFileFormat == THtmlFileFormat.MHtml && UseContentId)
                            {
                                ImageCid = GetImageId();
                                FullImageLink = GetContentId(ImageCid);
                            }
                            else
                            {
                                FullImageLink = EncodeUrl(e.ImageLink);
                            }

                            real ColWidth = Img.Width / Images.Props.ImageResolution * FlexCelRender.DispMul;
                            real RowHeight = Img.Height / Images.Props.ImageResolution * FlexCelRender.DispMul;

                            ImageTag =
                                FixIe6TranspBegin(RowHeight, ColWidth, FullImageLink, e.SavedImageFormat) +

                                FormatStr("<img src='{0}' width='{1}' height='{2}' alt=\"{3}\" " +
                                "style='border:none; width:{4" + PointsFormat + ";height:{5" + PointsFormat + ";' {6}",
                                FullImageLink, Img.Width, Img.Height, EncodeAsHtml(html, e.AlternateText, TEnterStyle.Char10),
                                ColWidth, RowHeight, EndOfTag) +

                                FixIe6TranspEnd(e.SavedImageFormat);

                        }

                        if (FHtmlFileFormat == THtmlFileFormat.Html)
                        {
                            SaveImageEventArgs sie = new SaveImageEventArgs(xls, -1, null, e.ImageFile, e.ImageLink, e.AlternateText, e.SavedImageFormat, Img);
                            Parent.OnSaveImage(sie);

                            if (!sie.Processed)
                            {
                                if (SaveImages)
                                {
                                    SaveImage(Img, e.ImageStream, e.ImageFile, e.SavedImageFormat);
                                }
                            }
                        }
                        else
                        {
                            Cache.HeaderImages.Add(new THtmlHeaderImageCache(sectionloc, FullImageLink, IsAbsoluteUrl(e.ImageLink), e.SavedImageFormat, ImageCid));
                        }
                    }
                }
            }

            WriteLn(html, xls.GetPageHeaderOrFooterAsHtml(Left, ImageTag, 1, 1, HtmlVersion, html.Encoding, Cache.OnHtmlFont));
            WriteLn(html, "</div>");

        }


        #endregion
    }
    #endregion

    #region Cache
    #region ExportHtmlCache
    internal class TExportHtmlCache
    {
        internal THtmlImageCacheList Images; //this doesn't contain images generated on the fly.
        internal THtmlCellImageCacheList CellImages; //This contains only images created on the fly, like images for vertical text.
        internal THtmlHeaderImageCacheList HeaderImages;
        internal THyperLinkCache Hyperlinks;
        internal TXFFormatCache Formats;
        internal TUsedFormatsCache UsedFormats;
        internal TCssInformation CssInfo;
        internal TSheetSelector SheetSelector;
        internal IHtmlFontEvent OnHtmlFont;

        internal TMergedCellsInColumn MergedCellsInColumn;
        internal TNamedRangeCache Names;

        public TExportHtmlCache(TCssInformation aCssInfo, TSheetSelector aSheetSelector, IHtmlFontEvent aOnHtmlFont)
        {
            SheetSelector = aSheetSelector;
            Images = new THtmlImageCacheList();
            CellImages = new THtmlCellImageCacheList();
            HeaderImages = new THtmlHeaderImageCacheList();
            Hyperlinks = new THyperLinkCache();
            Formats = new TXFFormatCache();
            OnHtmlFont = aOnHtmlFont;

            if (aCssInfo == null)
            {
                CssInfo = new TCssInformation(null, null);
            }
            else
            {
                CssInfo = aCssInfo;
            }
            UsedFormats = CssInfo.UsedFormats;
        }
    }
    #endregion

    #region THTMLImageCache
    internal class THtmlImageCache
    {
        internal int r1, c1, r2, c2;
        internal string Url;
        internal bool UrlIsAbsolute;
        internal int ZOrder;
        internal Size SizePixels;
        internal RectangleF Dimensions;
        internal PointF Origin;
        internal string AltText;
        internal string HyperLink;
        internal int PictureId;
        internal THtmlImageFormat SavedImageFormat;
        internal string UniqueId;

        internal THtmlImageCache(int ar1, int ac1, int ar2, int ac2, string aUrl, bool aUrlIsAbsolute, int aZOrder, Size aSizePixels, PointF aOrigin, RectangleF aDimensions, string aAltText, string aHyperLink,
            int aPictureId, THtmlImageFormat aSavedImageFormat, string aUniqueId)
        {
            r1 = ar1; c1 = ac1; r2 = ar2; c2 = ac2;
            Url = aUrl;
            UrlIsAbsolute = aUrlIsAbsolute;
            ZOrder = aZOrder;
            SizePixels = aSizePixels;
            Dimensions = aDimensions;
            Origin = aOrigin;
            AltText = aAltText;
            HyperLink = aHyperLink;
            PictureId = aPictureId;
            SavedImageFormat = aSavedImageFormat;
            UniqueId = aUniqueId;
        }
    }

#if (FRAMEWORK20)
    internal class THtmlImageCacheList : Dictionary<long, THtmlImageCache[]>
    {
#else
	internal class THtmlImageCacheList: Hashtable
	{
		private bool TryGetValue(long key, out THtmlImageCache[] Result)
		{
			Result = (THtmlImageCache[])this[key];	
			return Result != null;
		}
#endif

        public void Add(int r, int c, THtmlImageCache ImgData)
        {
            long key = FlxHash.MakeHash(r, c);

            THtmlImageCache[] CellImages;
            if (TryGetValue(key, out CellImages))
            {
                //This is not efficient, but normally we will not have more than one image in a cell. So for a normal case, having a 
                //simple array is much faster and uses less memory than having an arraylist. In general this should perform better than having an arraylist
                //to store multiple images in a cell.
                THtmlImageCache[] NewCellImages = new THtmlImageCache[CellImages.Length + 1];
                Array.Copy(CellImages, 0, NewCellImages, 0, CellImages.Length);
                NewCellImages[CellImages.Length] = ImgData;
                CellImages = NewCellImages;
            }
            else
            {
                CellImages = new THtmlImageCache[1];
                CellImages[0] = ImgData;
            }

            this[key] = CellImages;
        }

        public bool TryGetValue(int row, int col, int RowSpan, int ColSpan, out THtmlImageCache[,][] ResultValue)
        {
            ResultValue = null;
            for (int r = 0; r < RowSpan; r++)
            {
                for (int c = 0; c < ColSpan; c++)
                {
                    THtmlImageCache[] RCell;
                    if (TryGetValue(FlxHash.MakeHash(row + r, col + c), out RCell))
                    {
                        if (ResultValue == null)
                        {
                            ResultValue = new THtmlImageCache[RowSpan, ColSpan][];
                        }
                        ResultValue[r, c] = RCell;
                    }
                }
            }
            return ResultValue != null;
        }
    }
    #endregion

    #region THTMLCellImageCache
    internal class THtmlCellImageCache
    {
        internal int Row;
        internal int Col;
        internal string Url;
        internal string UniqueId;
        internal bool UrlIsAbsolute;
        internal THtmlImageFormat SavedImageFormat;

        internal THtmlCellImageCache(int aRow, int aCol, string aUrl, bool aUrlIsAbsolute, THtmlImageFormat aSavedImageFormat, string aUniqueId)
        {
            Row = aRow;
            Col = aCol;
            Url = aUrl;
            UrlIsAbsolute = aUrlIsAbsolute;
            SavedImageFormat = aSavedImageFormat;
            UniqueId = aUniqueId;
        }
    }

#if (FRAMEWORK20)
    internal class THtmlCellImageCacheList : List<THtmlCellImageCache>
    {
#else
	internal class THtmlCellImageCacheList: ArrayList
	{
		public new THtmlCellImageCache this[int index]
		{
			get
			{
				return (THtmlCellImageCache)base[index];
			}
			set
			{
				base[index] = value;
			}
		}
#endif
    }
    #endregion

    #region THTMLHeaderImageCache
    internal class THtmlHeaderImageCache
    {
        internal THeaderAndFooterPos Section;
        internal string Url;
        internal string UniqueId;
        internal bool UrlIsAbsolute;
        internal THtmlImageFormat SavedImageFormat;

        internal THtmlHeaderImageCache(THeaderAndFooterPos aSection, string aUrl, bool aUrlIsAbsolute, THtmlImageFormat aSavedImageFormat, string aUniqueId)
        {
            Section = aSection;
            Url = aUrl;
            UrlIsAbsolute = aUrlIsAbsolute;
            SavedImageFormat = aSavedImageFormat;
            UniqueId = aUniqueId;
        }
    }

#if (FRAMEWORK20)
    internal class THtmlHeaderImageCacheList : List<THtmlHeaderImageCache>
    {
#else
	internal class THtmlHeaderImageCacheList: ArrayList
	{
		public new THtmlHeaderImageCache this[int index]
		{
			get
			{
				return (THtmlHeaderImageCache)base[index];
			}
			set
			{
				base[index] = value;
			}
		}
#endif
    }
    #endregion

    #region Hyperlink Cache
    internal class THyperLinkCache : Dictionary<long, int>
    {
        internal void CacheHyperLinks(ExcelFile xls, TXlsCellRange PrintRange)
        {
            for (int i = xls.HyperLinkCount; i > 0; i--)
            {
                TXlsCellRange cr = xls.GetHyperLinkCellRange(i);
                if (cr.Top >= PrintRange.Top && cr.Top <= PrintRange.Bottom && cr.Left >= PrintRange.Left && cr.Left <= PrintRange.Right)
                {
                    this[FlxHash.MakeHash(cr.Top, cr.Left)] = i;
                }
            }
        }

    }
    #endregion

    #region Name cache
    internal class TNamedRangeCache : Dictionary<TOneCellRef, TXlsNamedRange>
    {
        internal void CacheNames(ExcelFile xls, TXlsCellRange PrintRange)
        {
            for (int i = xls.NamedRangeCount; i > 0; i--)
            {
                TXlsNamedRange nr = xls.GetNamedRange(i);
                if (nr.SheetIndex == xls.ActiveSheet && !nr.BuiltIn &&
                    nr.Top >= PrintRange.Top && nr.Top <= PrintRange.Bottom && nr.Left >= PrintRange.Left && nr.Left <= PrintRange.Right)
                {
                    this[new TOneCellRef(nr.Top, nr.Left)] = nr;
                }
            }
        }
    }
    #endregion

    #region Used Formats Cache
    internal class TUsedFormatsCache
    {
        Dictionary<TFlxFormat, TCachedFormat> Formats;
        Dictionary<TFlxFont, string> Fonts;
        Dictionary<string, int> UniqueFormats;

        internal TUsedFormatsCache()
        {
            Formats = new Dictionary<TFlxFormat, TCachedFormat>(new TUsedFormatComparer());
            Fonts = new Dictionary<TFlxFont, string>();
            UniqueFormats = new Dictionary<string, int>();
        }

        public void Add(ExcelFile xls, IHtmlFontEvent OnHtmlFont, TFlxFormat cfmt)
        {
            if (Formats.ContainsKey(cfmt)) return;

            string fontStr;
            if (!Fonts.TryGetValue(cfmt.Font, out fontStr))
            {
                string FontName = cfmt.Font.Name;
                string FinalFontName = FontName;
                if (FontName != null && FontName.IndexOf(" ") >= 0) FinalFontName = "'" + FontName + "'";  //font names with spaces must be quoted.

                HtmlFontEventArgs e = new HtmlFontEventArgs(xls, cfmt.Font, FinalFontName);
                if (OnHtmlFont != null) OnHtmlFont.DoHtmlFont(e);
                fontStr = e.FontFamily;
                Fonts.Add(cfmt.Font, fontStr);
            }

            TCachedFormat cafmt = new TCachedFormat(xls, cfmt, UniqueFormats.Count, fontStr);

            int existingfmt;
            if (UniqueFormats.TryGetValue(cafmt.Fmt, out existingfmt))
            {
                cafmt.Id = existingfmt;
            }
            else
            {
                UniqueFormats.Add(cafmt.Fmt, cafmt.Id);
            }
            Formats.Add(cfmt, cafmt);
        }

        public TCachedFormat this[TFlxFormat fmt]
        {
            get
            {
                return Formats[fmt];
            }
        }

        internal bool TryGetValue(TFlxFormat fmt, out TCachedFormat cfmt)
        {
            return Formats.TryGetValue(fmt, out cfmt);
        }

        internal int UniqueCount
        {
            get
            {
                return UniqueFormats.Count;
            }
        }

        public IEnumerable<string> Values { get { return UniqueFormats.Keys; } }

        internal int GetId(string fmt)
        {
            return UniqueFormats[fmt];
        }
    }

    internal class TUsedFormatComparer : IEqualityComparer<TFlxFormat>
    {
        #region IEqualityComparer<TFlxFormat> Members

        public bool Equals(TFlxFormat x, TFlxFormat y)
        {
            if (x.FillPattern.Pattern == TFlxPatternStyle.Solid)
            {
                if (y.FillPattern.Pattern != TFlxPatternStyle.Solid) return false;
                if (x.FillPattern.FgColor != y.FillPattern.FgColor) return false;
            }
            else
            {
                if (y.FillPattern.Pattern == TFlxPatternStyle.Solid) return false;
                if (x.FillPattern.BgColor != y.FillPattern.BgColor) return false;
            }


            if (x.Font.Color != y.Font.Color) return false;

            if (x.Font.Size20 != y.Font.Size20) return false;
            if (x.Font.Style != y.Font.Style) return false;

            if (x.Font.Name != y.Font.Name) return false;
            if (x.HAlignment != y.HAlignment) return false;
            if (x.VAlignment != y.VAlignment) return false;

            if (TCachedFormat.Wraps(x) != TCachedFormat.Wraps(y)) return false;

            return true;
        }

        public int GetHashCode(TFlxFormat x)
        {
            return HashCoder.GetHashObj(
                x.FillPattern.Pattern == TFlxPatternStyle.Solid ? x.FillPattern.FgColor : x.FillPattern.BgColor,

            x.Font.Color,
            x.Font.Size20,
            ((int)x.Font.Style),
            x.Font.Name,
            (int)x.HAlignment,
            (int)x.VAlignment,

            TCachedFormat.Wraps(x));
        }

        #endregion
    }


    internal struct TCachedFormat : IComparable
    {
        internal int Id;
        internal string Fmt;

        public TCachedFormat(ExcelFile xls, TFlxFormat fmt, int aId, string fontStr)
        {
            Id = aId;
            Fmt = GetCss(xls, fmt, fontStr);
        }

        internal static bool Wraps(TFlxFormat fmt)
        {
            return fmt.WrapText || fmt.HAlignment == THFlxAlignment.justify || fmt.VAlignment == TVFlxAlignment.justify;
        }

        internal static string GetCss(ExcelFile xls, TFlxFormat fmt, string FontFamily)
        {
            StringBuilder css = new StringBuilder();

            //if (fmt.FillPattern.Pattern != TFlxPatternStyle.None)  //None patterns should also be output.
            {
                if (fmt.FillPattern.Pattern == TFlxPatternStyle.Solid)
                {
                    WriteLn(css, "  background-color:" + THtmlEngine.GetColor(xls, fmt.FillPattern.FgColor, Color.White) + ";");
                }
                else
                {
                    WriteLn(css, "  background-color:" + THtmlEngine.GetColor(xls, fmt.FillPattern.BgColor, Color.White) + ";");
                }
            }

            WriteLn(css, "  color:" + THtmlEngine.GetColor(xls, fmt.Font.Color, Color.Black) + ";");

            css.Append("  font-size:");
            css.Append((fmt.Font.Size20 / 20.0).ToString("0.#", CultureInfo.InvariantCulture));
            WriteLn(css, "pt;");

            string StrBold = (fmt.Font.Style & TFlxFontStyles.Bold) != 0 ? "bold" : "normal";
            WriteLn(css, "  font-weight:" + StrBold + ";");

            string StrItalic = (fmt.Font.Style & TFlxFontStyles.Italic) != 0 ? "italic" : "normal";
            WriteLn(css, "  font-style:" + StrItalic + ";");

            /*if (fmt.Font.Underline == TFlxUnderline.Single || fmt.Font.Underline == TFlxUnderline.SingleAccounting)
            {
                WriteLn(css, "  text-underline-style:single;");
            }
            if (fmt.Font.Underline == TFlxUnderline.Double || fmt.Font.Underline == TFlxUnderline.DoubleAccounting)
            {
                WriteLn(css, "  text-underline-style:double;");
            }*/

            WriteLn(css, "  font-family:" + FontFamily + ";");

            string StrHAlign = "left";
            switch (fmt.HAlignment)
            {
                case THFlxAlignment.general:
                    break;
                case THFlxAlignment.left:
                    StrHAlign = "left";
                    break;
                case THFlxAlignment.center:
                    StrHAlign = "center";
                    break;
                case THFlxAlignment.right:
                    StrHAlign = "right";
                    break;
                case THFlxAlignment.fill:
                case THFlxAlignment.justify:
                case THFlxAlignment.center_across_selection:
                case THFlxAlignment.distributed:
                    StrHAlign = "justify";
                    break;
            }
            WriteLn(css, "  text-align:" + StrHAlign + ";");

            string StrVAlign = "top";
            switch (fmt.VAlignment)
            {
                case TVFlxAlignment.top:
                    StrVAlign = "top";
                    break;
                case TVFlxAlignment.center:
                    StrVAlign = "middle";
                    break;
                case TVFlxAlignment.bottom:
                    StrVAlign = "bottom";
                    break;
                case TVFlxAlignment.justify:
                    break;
                case TVFlxAlignment.distributed:
                    break;
                default:
                    break;
            }
            WriteLn(css, "  vertical-align:" + StrVAlign + ";");
            //WriteLn(css, "  overflow:hidden;"); //not here... if a row has overflow:hidden, firefox will not show correctly cells with rowspan > 1
            //WriteLn(css, "  clIm: auto;");

            string StrWrap = Wraps(fmt) ? "normal" : "nowrap";
            WriteLn(css, "  white-space:" + StrWrap + ";");
            WriteLn(css, " }");

            return css.ToString();

        }

        private static void WriteLn(StringBuilder css, string s)
        {
            css.Append(s);
            css.Append("\r\n");
        }

    #endregion

        #region IComparable Members

        public int CompareTo(object obj)
        {
            if (!(obj is TCachedFormat)) return -1;
            TCachedFormat o = (TCachedFormat)obj;
            return Fmt.CompareTo(o.Fmt);
        }

        #endregion
    }

    #endregion

    #region Merged cells by column cache
    class TMergedCellsInColumn
    {
#if (FRAMEWORK20)
        internal Dictionary<int, TColumnRangeData> Columns;
#else
		internal Hashtable Columns;
#endif


        internal TMergedCellsInColumn()
        {
#if (FRAMEWORK20)
            Columns = new Dictionary<int, TColumnRangeData>();
#else
            Columns = new Hashtable();
#endif
        }

        internal void Add(int r1, int c1, int r2, int c2)
        {
            if (r2 < r1 || c2 < c1) return;
            if (r1 == r2 && c1 == c2) return; //don't add single cell ranges.
            TOneRangeData RangeToAdd = new TOneRangeData(r1, c1, r2, c2);
            MixAndAddRange(RangeToAdd);
        }

        private bool TryGetValue(int c, out TColumnRangeData col)
        {
#if (FRAMEWORK20)
            return Columns.TryGetValue(c, out col);
#else
			col = Columns[c] as TColumnRangeData;
			return col != null;
#endif
        }

        private void MixAndAddRange(TOneRangeData RangeToAdd)
        {
            TOneRangeData ControlRange;
            do
            {
                ControlRange = RangeToAdd;
                for (int c = RangeToAdd.c1; c <= RangeToAdd.c2; c++)
                {
                    TColumnRangeData col;
                    if (!TryGetValue(c, out col))
                    {
                        col = new TColumnRangeData();
                        Columns.Add(c, col);
                    }
                    col.MixAndAddRanges(ref RangeToAdd);
                }
            } while (ControlRange != RangeToAdd);
        }

        internal TXlsCellRange GetRange(int r, int c)
        {
            TColumnRangeData Col;
            if (!TryGetValue(c, out Col)) return new TXlsCellRange(r, c, r, c);
            return Col.GetRange(r, c);
        }
    }

    class TColumnRangeData
    {
#if (FRAMEWORK20)
        private List<TOneRangeData> FList;
#else
		private ArrayList FList;
#endif

        internal TColumnRangeData()
        {
#if (FRAMEWORK20)
            FList = new List<TOneRangeData>();
#else
			FList = new ArrayList();
#endif
        }

        internal void MixAndAddRanges(ref TOneRangeData range)
        {
            if (FList.Count == 0)//  this avoids issues when ~index-1 < 0.
            {
                FList.Add(range);
                return;
            }

            int InsertIndex = -1;
            int index = FList.BinarySearch(range);

            if (index < 0)
            {
                InsertIndex = ~index;
                index = InsertIndex - 1; //when the item was not found, we still need to check the last item's r2 doesn't include this one r1.
                if (index < 0) index = 0;
            }

            int Added = -1;
            int k = index;
            while (k < FList.Count)
            {
                TOneRangeData FoundRange = (TOneRangeData)FList[k];

                if (FoundRange.r1 <= range.r2 && FoundRange.r2 >= range.r1)
                {
                    range.r1 = Math.Min(FoundRange.r1, range.r1);
                    range.r2 = Math.Max(FoundRange.r2, range.r2);
                    range.c1 = Math.Min(FoundRange.c1, range.c1);
                    range.c2 = Math.Max(FoundRange.c2, range.c2);
                    if (Added < 0) Added = k;
                    FList[Added] = range;

                    if (Added == k) k++;
                    else FList.RemoveAt(k);
                }
                else
                {
                    if (k == index) k++;  //it might happen that the block above didn't contain this one (BlockAbove_R2 < Block_R1), but still BlockBelow_R1 < Block_R2.
                    else
                    {
                        if (Added < 0) FList.Insert(InsertIndex, range);
                        return;
                    }
                }
            }
            if (Added < 0) FList.Insert(InsertIndex, range); //InsertIndex must not be -1, since if the object was found, it should have exited in the loop above.
        }

        internal TXlsCellRange GetRange(int r, int c)
        {
            TOneRangeData range = new TOneRangeData(r, 1, r, 1);
            int index = FList.BinarySearch(range);
            if (index < 0)
            {
                index = ~index - 1; //when the item was not found, we still need to check the last item's r2 doesn't include this one r1.
                if (index < 0) index = 0;
            }

            for (int i = index; i <= index + 1; i++)
            {
                TXlsCellRange Result = CheckRangeAt(r, i);
                if (Result != null) return Result;
            }

            return new TXlsCellRange(r, c, r, c);
        }

        private TXlsCellRange CheckRangeAt(int r, int index)
        {
            if (index >= FList.Count) return null;
            TOneRangeData FoundRange = (TOneRangeData)FList[index];
            if (FoundRange.r1 <= r && FoundRange.r2 >= r) return new TXlsCellRange(FoundRange.r1, FoundRange.c1, FoundRange.r2, FoundRange.c2);
            return null;
        }
    }


    struct TOneRangeData : IComparable
    {
        internal int r1;
        internal int r2;
        internal int c1;
        internal int c2;

        public TOneRangeData(int aR1, int aC1, int aR2, int aC2)
        {
            r1 = aR1;
            c1 = aC1;
            r2 = aR2;
            c2 = aC2;
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            if (!(obj is TOneRangeData)) return -1;
            TOneRangeData o2 = (TOneRangeData)obj;

            return r1.CompareTo(o2.r1);  //This holds ranges that don't overlap, so we can just compare the starting row.
        }

        public override bool Equals(object obj)
        {
            if (!(obj is TOneRangeData)) return false;
            TOneRangeData o2 = (TOneRangeData)obj;
            return (r1 == o2.r1 && r2 == o2.r2 && c1 == o2.c1 && c2 == o2.c2);
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHash(r1, c1, r2, c2);
        }

        public static bool operator ==(TOneRangeData o1, TOneRangeData o2)
        {
            return o1.Equals(o2);
        }

        public static bool operator !=(TOneRangeData o1, TOneRangeData o2)
        {
            return !o1.Equals(o2);
        }

        #endregion
    }
    #endregion

    #region TImageProps
    internal class TImageProps : ICloneable
    {
        internal real ImageResolution;
        internal SmoothingMode SmoothingMode;
        internal InterpolationMode InterpolationMode;
        internal bool AntiAliased;
        internal TImageNaming ImageNaming;
        internal Color ImageBackground;

        internal TImageProps()
        {
            ImageBackground = ColorUtil.Empty;
        }


        #region ICloneable Members

        public object Clone()
        {
            TImageProps Result = new TImageProps();
            Result.ImageResolution = ImageResolution;
            Result.SmoothingMode = SmoothingMode;
            Result.InterpolationMode = InterpolationMode;
            Result.AntiAliased = AntiAliased;
            Result.ImageNaming = ImageNaming;
            Result.ImageBackground = ImageBackground;
            return Result;
        }

        #endregion
    }
    #endregion

    #region TImageInformation
    internal class TImageInformation
    {
        internal TImageProps Props;
        internal THtmlImageFormat SavedImagesFormat;

        internal TImageInformation(TImageProps aProps, THtmlImageFormat aSavedImagesFormat)
        {
            Props = aProps;
            SavedImagesFormat = aSavedImagesFormat;
        }
    }
    #endregion

    #region TGeneratedFiles
    /// <summary>
    /// An object containing all the files generated in the export.
    /// </summary>
    public class TGeneratedFiles
    {
        #region Privates
        private List<string> FHtmlFiles;
        private List<string> FImageFiles;
        private List<string> FCssFiles;
        #endregion

        internal TGeneratedFiles()
        {
            Clear();
        }

        #region public properties
        /// <summary>
        /// Name of the html files generated.
        /// </summary>
        public string[] GetHtmlFiles()
        {
            return FHtmlFiles.ToArray();
        }

        /// <summary>
        /// Name of the image files generated.
        /// </summary>
        public string[] GetImageFiles()
        {
            return FImageFiles.ToArray();
        }

        /// <summary>
        /// Name of the css files generated.
        /// </summary>
        public string[] GetCssFiles()
        {
            return FCssFiles.ToArray();
        }
        #endregion

        /// <summary>
        /// Clears all the files in the object.
        /// </summary>
        public void Clear()
        {
            FHtmlFiles = new List<string>();
            FImageFiles = new List<string>();
            FCssFiles = new List<string>(1);
        }

        internal void AddHtml(string FileName)
        {
            FHtmlFiles.Add(FileName);
        }

        internal void AddImage(string Image)
        {
            FImageFiles.Add(Image);
        }

        internal void AddCss(string Css)
        {
            FCssFiles.Add(Css);
        }

        internal void Add(string FileName, THtmlFileType FileType)
        {
            switch (FileType)
            {
                case THtmlFileType.Html:
                    AddHtml(FileName);
                    break;
                case THtmlFileType.Css:
                    AddCss(FileName);
                    break;
                case THtmlFileType.Image:
                    AddImage(FileName);
                    break;
            }
        }
    }

    #endregion

    #region THtmlFileType
    internal enum THtmlFileType
    {
        Html,
        Css,
        Image
    }
    #endregion

    #region ImageContainer
    internal class ImageContainer : IDisposable
    {
        private bool NeedsDispose;
        internal Image Img;

        internal ImageContainer(Image aImg, bool aNeedsDispose)
        {
            Img = aImg;
            NeedsDispose = aNeedsDispose;
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (NeedsDispose && Img != null) Img.Dispose();
            Img = null;
            GC.SuppressFinalize(this);
        }

        #endregion

    }

    #endregion

    #region SaveAndHandleSharingViolation
    internal abstract class TSaveAndHandleSharingViolation
    {

        internal void Save(bool IgnoreSharingViolations, bool AllowOverwritingFiles, string FileName)
        {
            //This is a problematic method, because C# does not allow direct access to HResult in IOException, and
            //also IOException is too generic to just catch it and assume the error was because the file was locked.
            //Even when this is on the examples in msdn: http://msdn2.microsoft.com/en-us/library/system.io.ioexception.aspx
            //So we will be adding 2 more checks: First a boolean flag that will be true if the error was not when opening the file, and
            //we will also check for the HResult, by using GetHRForException. If both things indicate a sharing violation we will just exit,
            //otherwise rethrow the exception.
            bool Opened = false;
            try
            {
                FileMode fm = FileMode.CreateNew;
                if (AllowOverwritingFiles) fm = FileMode.Create;

                using (FileStream f = new FileStream(FileName, fm, FileAccess.Write, FileShare.None))
                {
                    Opened = true;
                    InternalSave(f);
                }
            }
            catch (IOException ex)
            {
                if (!IgnoreSharingViolations || !AllowOverwritingFiles) throw;
                if (Opened) throw;
                if (!FlxUtils.HasUnamanagedPermissions()) throw;
                if (CheckRethrow(ex)) throw;
                if (FlexCelTrace.HasListeners) FlexCelTrace.Write(new THtmlSaveSharingViolationError(ex.Message, FileName));
            }

        }

        [MethodImpl(MethodImplOptions.NoInlining)]
#if (FRAMEWORK40)
        [SecuritySafeCritical]
#endif
        internal static bool CheckRethrow(IOException ex)
        {
#if (!FULLYMANAGED && FRAMEWORK20)

            int HResult = 0;
            try
            {
                HResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
            }
            catch (SecurityException)
            {
            }

            if (HResult != unchecked((int)0x80070020)) return true;
#else

			//Sadly .net 1.0 does not return the real error code in hresult, it returns a bogus COR_E_IO which has the value 0x80131620
			//So we cannot find out if it was a sharing violation or not ,and we must assume it was.

#endif

            return false;
        }

        protected abstract void InternalSave(FileStream f);
    }

    internal class TSaveImageSV : TSaveAndHandleSharingViolation
    {
        private Image Img;
        private THtmlImageFormat ImgFormat;

        internal TSaveImageSV(Image aImg, THtmlImageFormat aImgFormat)
        {
            Img = aImg;
            ImgFormat = aImgFormat;
        }

        protected override void InternalSave(FileStream f)
        {
            Img.Save(f, THtmlEngine.GetImageFormat(ImgFormat));
        }

    }
    #endregion
}
