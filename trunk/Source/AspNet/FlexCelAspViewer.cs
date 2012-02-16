using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using FlexCel.Core;
using FlexCel.Render;
using System.Web.UI.HtmlControls;
using System.Web.Security;
using System.IO;
using System.Net;
using System.Threading;

namespace FlexCel.AspNet
{
    #region Viewer
    /// <summary>
    /// An ASP.NET component that can be used to embed an Excel file as HTML in a WebForm.
    /// </summary>
    [ToolboxData("<{0}:FlexCelAspViewer runat=server></{0}:FlexCelAspViewer>")]
    public class FlexCelAspViewer : WebControl
    {
        #region Privates
        private FlexCelAspExport FHtmlExport;
        private THtmlSheetExport FSheetExport;
        private string FSheetSeparator;

        private string FRelativeImagePath;

        TPartialExportState FExportState;

        TImageExportMode FImageExportMode;

        string FImageHandlerName;
        string FImageParameterName;

        int FImageTimeout;
        private static DateTime LastImageCleaned;
        private int FMaxTemporaryImages;
        private static object TimeoutAccess = new object();

        #endregion

        #region Properties
        /// <summary>
        /// Component that will do the actual conversion from an Excel file to HTML. Most of the export settings should be changed in this component.
        /// </summary>
        [Category("Engine"), Browsable(false),   //If we make it browsable, it will show in the designer, but will not save any changes.
        Description("Component that will do the actual conversion from an Excel file to HTML. Most of the export settings should be changed in this component.")]
        public FlexCelAspExport HtmlExport { get { return FHtmlExport; } }


        /// <summary>
        /// Defines how many sheets will be exported.
        /// </summary>
        [Category("Behavior"),
        Description("Defines how many sheets will be exported."),
        DefaultValue(THtmlSheetExport.ActiveSheet)]
        public THtmlSheetExport SheetExport { get { return FSheetExport; } set { FSheetExport = value; } }

        /// <summary>
        /// Defines a separator between multiple sheets when exporting all visible sheets. You can use for example &lt;hr&gt; here to write an
        /// horizontal ruler between sheets.
        /// </summary>
        [Category("Behavior"),
        Description("Defines a separator between multiple sheets when exporting all visible sheets. You can use for example &lt;hr&gt; here to write an horizontal ruler between sheets.")]
        public string SheetSeparator { get { return FSheetSeparator; } set { FSheetSeparator = value; } }


        /// <summary>
        /// Path where the images will be stored, relative to the path where the page is in.
        /// Try to use a different folder for the images, so it is easier to delete them later.
        /// </summary>
        [Browsable(true), Category("Images"),
        Description("Path where the images will be stored, relative to the path where the page is in. Try to use a different folder for the images, so it is easier to delete them later."),
        DefaultValue("images")]
        public string RelativeImagePath { get { return FRelativeImagePath; } set { FRelativeImagePath = value; } }

        /// <summary>
        /// How the images will be served to the browser. See the PDF documentation on creating HTML files for more information.
        /// </summary>
        [Browsable(true), Category("Images"),
        Description("How the images will be served to the browser. See the PDF documentation on creating HTML files for more information."),
        DefaultValue(TImageExportMode.TemporaryFiles)]
        public TImageExportMode ImageExportMode { get { return FImageExportMode; } set { FImageExportMode = value; } }

        /// <summary>
        /// Time in minutes that temporary images will live in the server before being deleted. 
        /// Temporary images older than "current date - imageTimeOut" will be removed each time a new call to this component is made.
        /// Set this property to 0 or a negative value to not delete any image, if you wish to do the cleanup yourself using a server script.
        /// See also <see cref="MaxTemporaryImages"/>
        /// </summary>
        [Browsable(true), Category("Images"),
        Description("Time in minutes that temporary images will live in the server before being deleted. Set it to 0 or a negative value to not delete any temporary image."),
        DefaultValue(15)]
        public int ImageTimeout { get { return FImageTimeout; } set { FImageTimeout = value; } }

        /// <summary>
        /// Maximum number of temporary images in the images folder. Set it to 0 or a negative value if you want unlimited images.
        /// Use this variable to avoid Denial of Service conditions. For example, a malicious user could keep continuously refreshing the webpage,
        /// without waiting for the page to load. As all images are created each time he refreshes, but they are never deleted (because they are never requested
        /// back) the disk could fill very fast, not giving time to the <see cref="ImageTimeout"/> timespan to happen, and crashing the server.
        /// <para>Note that this number is approximate, if for example the maximum is 5000, you could get 5200 images in a peak time. It is not guaranteed that the
        /// maximum images will be 5000, just that it will not grow much more than that.</para>
        /// </summary>
        [Browsable(true), Category("Images"),
        Description("Maximum number of temporary images in the images folder."),
        DefaultValue(5000)]
        public int MaxTemporaryImages { get { return FMaxTemporaryImages; } set { FMaxTemporaryImages = value; } }

        /// <summary>
        /// Name for the ImageHandler used to return images when <see cref="ImageExportMode"/> is not <see cref="TImageExportMode.TemporaryFiles"/>.
        /// You need to add this name in web.config in your app.
        /// </summary>
        [Browsable(true), Category("Images"),
        Description("Name for the ImageHandler used to return images when ImageExportMode is not TImageExportMode.TemporaryFiles"),
        DefaultValue("flexcelviewer.ashx")]
        public string ImageHandlerName { get { return FImageHandlerName; } set { FImageHandlerName = value; } }

        /// <summary>
        /// Name for the parameter used to return the image in the URL. name when <see cref="ImageExportMode"/> is <see cref="TImageExportMode.UniqueTemporaryFiles"/>.
        /// This parameter will appear in the url as: http://server/.../flexcelviewer.ashx?IMAGEPARAMETERNAME=Imagename
        /// </summary>
        [Browsable(true), Category("Images"),
        Description("Name for the parameter used to return the image in the URL. This parameter will appear in the url as: http://server/.../flexcelviewer.ashx?IMAGEPARAMETERNAME=Imagename"),
        DefaultValue("image")]
        public string ImageParameterName { get { return FImageParameterName; } set { FImageParameterName = value; } }

        #endregion

        #region Events
        /// <summary>
        /// Use this event to customize the links in the HTML file when using <see cref="TImageExportMode.CustomStorage"/>.
        /// </summary>
        [Category("Images"),
        Description("Use this event to customize the links in the HTML file when using TImageExportMode.CustomStorage.")]
        public event ImageLinkEventHandler ImageLink;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelAspViewer.
        /// </summary>
        /// <param name="e"></param>
        protected internal virtual void OnImageLink(ImageLinkEventArgs e)
        {
            if (ImageLink != null) ImageLink(this, e);
        }

        /// <summary>
        /// Use this event to save the images into other place. You will normally need to use this event when implementing your own Http Handler to
        /// return the images.
        /// </summary>
        [Category("Images"),
        Description("Use this event to save the images into other place.")]
        public event SaveImageEventHandler SaveImage;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelAspViewer.
        /// </summary>
        /// <param name="e"></param>
        protected internal virtual void OnSaveImage(SaveImageEventArgs e)
        {
            if (SaveImage != null) SaveImage(this, e);
        }

        internal bool HasSaveImageEvent
        {
            get { return SaveImage != null; }
        }

        /// <summary>
        /// Use this event to customize how a named range if exported to the HTML file. Note that for this event to be called,
        /// you first need to set <see cref="HtmlExport"/>.ExportNamedRanges = true. If you want to change the id that will be exported or
        /// exclude certain named from being exported, you can do so here.
        /// </summary>
        [Category("Named Ranges"),
        Description("Use this event to customize how a named range if exported to the HTML file.")]
        public event NamedRangeExportEventHandler NamedRangeExport;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelAspViewer.
        /// </summary>
        /// <param name="e"></param>
        protected internal virtual void OnNamedRangeExport(NamedRangeExportEventArgs e)
        {
            if (NamedRangeExport != null) NamedRangeExport(this, e);
        }

        #endregion

        #region Constructor
        /// <summary>
        /// Creates a new instance of FlexCelAspViewer.
        /// </summary>
        public FlexCelAspViewer()
            : base(HtmlTextWriterTag.Div)
        {
            FHtmlExport = new FlexCelAspExport(this);
            FHtmlExport.HtmlVersion = THtmlVersion.XHTML_10; //default in 2.0 and up. this component doesn't support 1.1
            SheetExport = THtmlSheetExport.ActiveSheet;
            ImageExportMode = TImageExportMode.TemporaryFiles;
            FRelativeImagePath = "images";
            ImageHandlerName = "flexcelviewer.ashx";
            ImageParameterName = "image";
            ImageTimeout = 15;
            MaxTemporaryImages = 5000;
        }
        #endregion

        #region Implementation
        /// <summary>
        /// This prerender method is overriden so it saves the CSS rules to the head part of the html file.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnPreRender(EventArgs e)
        {
            base.OnPreRender(e);

            if (HtmlExport == null || HtmlExport.Workbook == null) return;


            string HtmlFileName = HttpContext.Current.Request.PhysicalPath;

            DeleteTemporaryImages();


            FExportState = new TPartialExportState(null, null);
            HtmlExport.ClassPrefix = "flx_" + this.ID + "_";


            if (SheetExport == THtmlSheetExport.AllVisibleSheets)
            {
                int SaveActiveSheet = HtmlExport.Workbook.ActiveSheet;
                try
                {
                    for (int sheet = 1; sheet <= HtmlExport.Workbook.SheetCount; sheet++)
                    {
                        HtmlExport.Workbook.ActiveSheet = sheet;
                        if (HtmlExport.Workbook.SheetVisible != TXlsSheetVisible.Visible) continue;

                        HtmlExport.PartialExportAdd(FExportState, HtmlFileName, FRelativeImagePath, ImageExportMode != TImageExportMode.CustomStorage);
                    }
                }
                finally
                {
                    HtmlExport.Workbook.ActiveSheet = SaveActiveSheet;
                }
            }
            else
            {
                HtmlExport.PartialExportAdd(FExportState, HtmlFileName, FRelativeImagePath, ImageExportMode != TImageExportMode.CustomStorage);
            }

            StyleSheetControl StyleSheet = new StyleSheetControl(FExportState);
            this.Page.Header.Controls.Add(StyleSheet);
        }

        /// <summary>
        /// Method overriden to return the exported xls file as HTML.
        /// </summary>
        /// <param name="writer"></param>
        protected override void RenderContents(HtmlTextWriter writer)
        {
            if (HtmlExport == null)
            {
                writer.Write("No FlexCelHtmlExport engine selected.");
                return;
            }

            if (HtmlExport.Workbook == null)
            {
                writer.Write("No Excel Workbook is assigned to the FlexCelHtmlExport component.");
                return;
            }

            //In case it was changed by other similar component in pre-render.
            HtmlExport.ClassPrefix = "flx_" + this.ID + "_";

            for (int i = 1; i <= FExportState.BodyCount; i++)
            {
                if (i > 1 && SheetSeparator != null) writer.WriteLine(SheetSeparator);
                FExportState.SaveBody(writer, i, RelativeImagePath);
            }


        }


        private int CompareFileDates(FileInfo a, FileInfo b)
        {
            return b.CreationTime.CompareTo(a.CreationTime);
        }

        private void DeleteTemporaryImages()
        {
            if (ImageTimeout <= 0 && MaxTemporaryImages <= 0) return;
            TimeSpan Timeout = ImageTimeout <= 0 ? new TimeSpan() : new TimeSpan(0, ImageTimeout, 0);

            if (!Monitor.TryEnter(TimeoutAccess)) return; //Only one thread at a time might do the cleanup.
            try
            {
                DateTime Now = DateTime.Now;
                try
                {
                    DirectoryInfo di = new DirectoryInfo(Path.Combine(MapPathSecure("~"), RelativeImagePath));

                    FileInfo[] files = di.GetFiles("????????????????????????????????.*");

                    if (files == null) return;
                    if ((MaxTemporaryImages <= 0 || files.Length <= MaxTemporaryImages) &&
                        (ImageTimeout <= 0 || Now - LastImageCleaned < Timeout)) return;

                    Array.Sort<FileInfo>(files, CompareFileDates);


                    for (int i = files.Length - 1; i >= 0; i--)
                    {

                        FileInfo fi = files[i];

                        switch (fi.Extension.ToLowerInvariant())
                        {
                            case ".png":
                            case ".jpg":
                            case ".jpeg":
                            case ".gif":
                                if ((MaxTemporaryImages > 0 && i > MaxTemporaryImages) ||
                                    (ImageTimeout > 0 && Now - fi.CreationTime > Timeout)) File.Delete(fi.FullName);
                                else return;
                                break;
                        }
                    }
                }
                finally
                {
                    LastImageCleaned = Now;
                }
            }
            finally
            {
                Monitor.Exit(TimeoutAccess);
            }
        }


        #endregion
    }
    #endregion

    #region Supporting classes and enums
    /// <summary>
    /// Defines how many sheets will be exported.
    /// </summary>
    public enum THtmlSheetExport
    {
        /// <summary>
        /// Only ActiveSheet on the workbook will be exported.
        /// </summary>
        ActiveSheet,

        /// <summary>
        /// All visible sheets in the workbook will be exported, one after the other.
        /// </summary>
        AllVisibleSheets
    }

    /// <summary>
    /// How images on the Excel file will be stored in order to be served to the browser.
    /// </summary>
    public enum TImageExportMode
    {
        /// <summary>
        /// When this option is selected, images will be saved in a temporary directory (defined in <see cref="FlexCel.AspNet.FlexCelAspViewer.RelativeImagePath"/> )
        /// You can configure a timeout period with <see cref="FlexCel.AspNet.FlexCelAspViewer.ImageTimeout"/> after which the images will be deleted.
        /// </summary>
        TemporaryFiles,

        /// <summary>
        /// When this option is selected, images will be saved in a temporary directory (defined in <see cref="FlexCel.AspNet.FlexCelAspViewer.RelativeImagePath"/> )
        /// You can configure a timeout period with <see cref="FlexCel.AspNet.FlexCelAspViewer.ImageTimeout"/> after which the images will be deleted.
        /// Different from <see cref="TImageExportMode.TemporaryFiles"/>, this mode will delete the image once it is served. 
        /// This is an advantage from the point of view that images are cleaned faster (without waiting for the timeout), but might be
        /// a disadvantage if the user for example right clicks the image and selects "view image". As the image is deleted the first time it is served,
        /// the second time (when the user does "view image") it will return image not found. Also some browsers might ask for the image more than once.<br/>
        /// <b>Use this mode only in controlled environments, whenre you know the browsers being used and you can test them.</b> If you use this mode as for a public web server, some users might not see the images.
        /// </summary>
        UniqueTemporaryFiles,

        /// <summary>
        /// This is an advanced option that allows you to store the images somewhere else, for example a database. Links on the image will point to
        /// an HttpHandler that you need to implement. See the PDF documentation on creating HTML files for more information.
        /// </summary>
        CustomStorage

    }

    #region ImageLink Event Handler
    /// <summary>
    /// Arguments passed on <see cref="FlexCel.AspNet.FlexCelAspViewer.ImageLink"/>, 
    /// </summary>
    public class ImageLinkEventArgs : EventArgs
    {
        private readonly ExcelFile FWorkbook;

        private readonly int FObjectIndex;
        private readonly TShapeProperties FShapeProps;
        private string FImageLink;
        private readonly string FAlternateText;
        private readonly THtmlImageFormat FSavedImageFormat;


        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        /// <param name="aWorkbook">See <see cref="Workbook"/></param>
        /// <param name="aObjectIndex">See <see cref="ObjectIndex"/></param>
        /// <param name="aShapeProps">See <see cref="ShapeProps"/></param>
        /// <param name="aImageLink">See <see cref="ImageLink"/></param>
        /// <param name="aAlternateText">See <see cref="AlternateText"/></param>
        /// <param name="aSavedImageFormat">See <see cref="SavedImageFormat"/></param>
        public ImageLinkEventArgs(ExcelFile aWorkbook, int aObjectIndex, TShapeProperties aShapeProps, string aImageLink, string aAlternateText, THtmlImageFormat aSavedImageFormat)
        {
            FWorkbook = aWorkbook;
            FObjectIndex = aObjectIndex;
            FShapeProps = aShapeProps;
            FImageLink = aImageLink;
            FAlternateText = aAlternateText;
            FSavedImageFormat = aSavedImageFormat;
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
        /// Alternate text for the image, to show in the "ALT" tag when a browser can not display images.
        /// By default this is set to the text in the box "Alternative Text" in the web tab on the image properties.
        /// If no Alternative text is supplied in the file, the image name will be used here.
        /// </summary>
        public string AlternateText { get { return FAlternateText; } }

        /// <summary>
        /// File format in which the image is. 
        /// </summary>
        public THtmlImageFormat SavedImageFormat { get { return FSavedImageFormat; } }

        /// <summary>
        /// The link that will be inserted in the html file. Modify it to create your won link to the HttpHandler.
        /// </summary>
        public string ImageLink { get { return FImageLink; } set { FImageLink = value; } }
    }

    /// <summary>
    /// Delegate used to specify where to store the images on a page.
    /// </summary>
    public delegate void ImageLinkEventHandler(object sender, ImageLinkEventArgs e);
    #endregion

    #endregion

    #region FlexCelAspExport
    /// <summary>
    /// A FlexCelHtmlExport specialized to export pages with FlexCelAspViewer.
    /// This class is mainly for internal use, there is no need to use it directly.
    /// </summary>
    public class FlexCelAspExport : FlexCelHtmlExport
    {
        private FlexCelAspViewer Viewer;

        /// <summary>
        /// Constructs a new instance of FlexCelHtmlExport.
        /// </summary>
        /// <param name="aViewer">Viewer that this exporter will be serving.</param>
        public FlexCelAspExport(FlexCelAspViewer aViewer)
        {
            Viewer = aViewer;
            ImageNaming = TImageNaming.Guid;
        }

        /// <summary>
        /// Intercepts the original OnGetImageInformation event to provide the parameters for FlexCelAspExport.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnGetImageInformation(ImageInformationEventArgs e)
        {
            switch (Viewer.ImageExportMode)
            {
                case TImageExportMode.TemporaryFiles:
                    break;

                case TImageExportMode.UniqueTemporaryFiles:
                    e.ImageLink = Viewer.ImageHandlerName + "?" + Viewer.ImageParameterName + "=" + e.ImageLink;
                    break;

                case TImageExportMode.CustomStorage:
                    ImageLinkEventArgs eLink = new ImageLinkEventArgs(e.Workbook, e.ObjectIndex, e.ShapeProps, e.ImageLink, e.AlternateText, e.SavedImageFormat);
                    Viewer.OnImageLink(eLink);
                    e.ImageLink = eLink.ImageLink;
                    e.ImageFile = null;
                    break;
            }

            base.OnGetImageInformation(e);
        }

        /// <summary>
        /// Intercepts the original OnGetSaveImage event to provide the parameters for FlexCelAspExport.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnSaveImage(SaveImageEventArgs e)
        {
            Viewer.OnSaveImage(e);
        }

        /// <summary>
        /// Returns true if the viewer has a SaveImage event assigned.
        /// </summary>
        public override bool HasSaveImageEvent
        {
            get
            {
                return Viewer.HasSaveImageEvent;
            }
        }

        /// <summary>
        /// Intercepts the original OnNamedRangeExport event to provide the parameters for FlexCelAspExport.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnNamedRangeExport(NamedRangeExportEventArgs e)
        {
            Viewer.OnNamedRangeExport(e);
        }
    }
    #endregion

    #region StyleSheetControl
    internal class StyleSheetControl : HtmlControl
    {
        TPartialExportState ExportState;

        internal StyleSheetControl(TPartialExportState aExportState)
        {
            ExportState = aExportState;
        }

        protected override void Render(HtmlTextWriter writer)
        {
            ExportState.SaveRelevantHeaders(writer);
        }
    }
    #endregion

    #region ImageHandler

    /// <summary>
    /// This is a base image handler you can use to create your own.
    /// This code is adapted from here: http://www.hanselman.com/blog/PermaLink,guid,5c59d662-b250-4eb2-96e4-f274295bd52e.aspx
    /// We would like to thank Scott Hanselman for the  insightful article.
    /// </summary>
    /// <remarks>
    /// To use this handler, include the following lines in a Web.config file.
    /// &lt;configuration&gt;
    ///    &lt;system.web&gt;
    ///       &lt;httpHandlers&gt;
    ///          &lt;add verb="*" path="flexcelviewer.ashx" type="FlexCel.AspNet.FlexCelHtmlImageHandler,FlexCel.AspNet"/&gt;
    ///       &lt;/httpHandlers&gt;
    ///    &lt;/system.web&gt;
    /// &lt;/configuration&gt;
    /// <br/>
    /// You might need to change the name "flexcelviewer.ashx" if you change <see cref="FlexCelAspViewer.ImageHandlerName"/>. Both names need to be the same.
    /// </remarks>
    public class FlexCelHtmlImageHandler : IHttpHandler
    {
        #region IHttpHandler Members

        /// <summary>
        /// This instance can be reused.
        /// </summary>
        public bool IsReusable
        {
            get { return true; }
        }

        /// <summary>
        /// This method will validate the paramters, and then call <see cref="GetImage"/> to get the real data.
        /// When inheriting from this class, you normally only have to override <see cref="ValidateParameters"/> <see cref="GetImage"/>
        /// </summary>
        /// <param name="context"></param>
        public virtual void ProcessRequest(HttpContext context)
        {
            if (!ValidateParameters(context))
            {
                //Internal Server Error
                context.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                context.Response.End();
                return;
            }

            if (RequiresAuthentication && !context.User.Identity.IsAuthenticated)
            {
                //Forbidden
                context.Response.StatusCode = (int)HttpStatusCode.Forbidden;
                context.Response.End();
                return;
            }

            if (!ImageExists(context))
            {
                //not found
                context.Response.StatusCode = (int)HttpStatusCode.NotFound;
                context.Response.End();
                return;
            }

            if (!GetImage(context))
            {
                //Internal Server Error
                context.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                context.Response.End();
                return;
            }

            //Everything is ok.
        }

        /// <summary>
        /// Override this method to check the paramters are valid.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public virtual bool ValidateParameters(HttpContext context)
        {
            return false;
        }

        /// <summary>
        /// Override this property in a descendant class if you do not want to ask for authentication for this image handler.
        /// </summary>
        public virtual bool RequiresAuthentication
        {
            get { return true; }
        }

        /// <summary>
        /// Override this method to provide a 404 error if the image has been deleted or doesn't exists.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public virtual bool ImageExists(HttpContext context)
        {
            return false;
        }

        /// <summary>
        /// Override this method to return the image. <br/>
        /// <b>WARNING:</b> Whatever you do here, make <b>really</b> sure you return only the indented image.
        /// A NAIVE IMPLEMENTATION OF THIS METHOD MIGHT RESULT IN A HUGE SECURITY HOLE. An attacker could use an implementation
        /// that just returns a file to retrieve any file in the server.
        /// </summary>
        /// <param name="context"></param>
        /// <returns>True if the image was correctly served.</returns>
        public virtual bool GetImage(HttpContext context)
        {
            return false;
        }

        #endregion
    }


    #endregion

    #region Image handler for Unique temporary files
    /// <summary>
    /// The image handler used when you select <see cref="TImageExportMode.UniqueTemporaryFiles"/>
    /// </summary>
    /// <remarks>
    /// To use this handler, include the following lines in a Web.config file.
    /// &lt;configuration&gt;
    ///    &lt;system.web&gt;
    ///       &lt;httpHandlers&gt;
    ///          &lt;add verb="*" path="flexcelviewer.ashx" type="FlexCel.AspNet.UniqueTemporaryFilesImageHandler,FlexCel.AspNet"/&gt;
    ///       &lt;/httpHandlers&gt;
    ///    &lt;/system.web&gt;
    /// &lt;/configuration&gt;
    /// <br/>
    /// You might need to change the name "flexcelviewer.ashx" if you change <see cref="FlexCelAspViewer.ImageHandlerName"/>. Both names need to be the same.
    /// <para>
    /// Note that this handler requires authentication in order to display the images. If you want to have a handler that does not require authentication,
    /// you can derive a class from this one, and override the <see cref="FlexCelHtmlImageHandler.RequiresAuthentication"/> property to return false. After that, register this new
    /// class in web.config instead of this one.
    /// </para>
    /// </remarks>
    public class UniqueTemporaryFilesImageHandler : FlexCelHtmlImageHandler
    {
        /// <summary>
        /// This method will validate the image passed is a valid GUID.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public override bool ValidateParameters(HttpContext context)
        {
            string FullImageName = context.Request.QueryString["image"];
            if (FullImageName == null) return false;
            string ImageName = Path.GetFileName(FullImageName);
            if (ImageName.Length != 32 + 4) return false; //The image should be a guid + extension.
            try
            {
                Guid g = new Guid(Path.GetFileNameWithoutExtension(ImageName));
            }
            catch (OverflowException)
            {
                return false;
            }
            catch (FormatException)
            {
                return false;
            }

            string ext = Path.GetExtension(ImageName).ToLowerInvariant();
            if (ext != ".png" && ext != ".jpg" && ext != ".jpeg" && ext != ".gif") return false;

            try
            {
                ImageFileName(FullImageName, context); //This will verify image is in the app path.
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static string ImageFileName(string ImageName, HttpContext context)
        {
            return context.Request.MapPath(ImageName, null, false);
        }

        /// <summary>
        /// This method checks if the image still exists, and return false if it doesn't.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public override bool ImageExists(HttpContext context)
        {
            string ImageName = ImageFileName(context.Request.QueryString["image"], context);
            return File.Exists(ImageName);
        }

        /// <summary>
        /// This method returns the image from the file.
        /// </summary>
        /// <param name="context"></param>
        public override bool GetImage(HttpContext context)
        {
            //ImageName has already been validated in ValidateParams. We must be sure there is no way for someone forging 
            //URLs to access something they should not.
            string ImageName = ImageFileName(context.Request.QueryString["image"], context);

            context.Response.ContentType = GetContentType(Path.GetExtension(ImageName));
            context.Response.WriteFile(ImageName);
            context.Response.Flush();
            context.Response.Close();


            //Image has been served, delete it.
            File.Delete(ImageName);

            return true;
        }

        private static string GetContentType(string p)
        {
            switch (p.ToLowerInvariant())
            {
                case "gif": return "image/gif";
                case "jpg":
                case "jpeg":
                    return "image/jpeg";
            }
            return "image/png";

        }


    }

    #endregion

}
