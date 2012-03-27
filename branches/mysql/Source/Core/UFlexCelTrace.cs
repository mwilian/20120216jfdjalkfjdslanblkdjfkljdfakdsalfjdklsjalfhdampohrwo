using System;
using System.Diagnostics;

#if (WPF)
using System.Windows;
#else
using System.Drawing;
#endif
namespace FlexCel.Core
{
	#region FlexCelError
	/// <summary>
	/// An enumeration of all possible FlexCel non fatal errors that can be logged.
	/// </summary>
	public enum FlexCelError
	{
        /// <summary>
        /// This error should never happen.
        /// </summary>
        Undefined = 0,

		/// <summary>
		/// There are more page breaks in this page than the maximum allowed, and page breaks over this maximum are being ignored.
		/// You can control if you want a "real" exception here with the <see cref="ExcelFile.ErrorActions"/> and <see cref="FlexCel.Report.FlexCelReport.ErrorActions"/> properties.
		/// </summary>
		XlsTooManyPageBreaks = 0x100,

        /// <summary>
        /// 
        /// </summary>
        RowHeightTooBig = 0x101,

        /// <summary>
        /// A corrupt xlsx file can have invalid names. If <see cref="ExcelFile.ErrorActions"/> is set to the 
        /// correct value, those names will be imported as #REF! instead of rising an Exception.
        /// </summary>
        XlsxInvalidName = 0x102,

        /// <summary>
        /// The file has references to other parts that don't exist. For example, an image might be referenced in a sheet, 
        /// but the actual image data missing from the file. If <see cref="ExcelFile.ErrorActions"/> is set to the 
        /// correct value, those parts will be ignored.
        /// </summary>
        XlsxMissingPart = 0x103,

		/// <summary>
		/// There was a GDI+ error trying to draw or print a metafile. It will be rendered as a bitmap.
		/// </summary>
		RenderMetafile = 0x200,

		/// <summary>
		/// The image could not be rendered by GDI+.
		/// </summary>
		RenderCorruptImage = 0x201,

		/// <summary>
		/// An image could not be drawn. This error normally happens in .NET 1.1 if vj# is not installed.
		/// </summary>
		RenderErrorDrawingImage = 0x202,

		/// <summary>
		/// The font was not found in the system. A substitute font will be used for the rendering, and it might not look the same.
		/// <br/> If you are seeing this message in a server, make sure the server has the same fonts installed as the machine where you
		/// developed the application.
		/// </summary>
		PdfFontNotFound = 0x300,

		/// <summary>
		/// FlexCel is trying to render a character that is not in the font or the fallback fonts. This character will show as an empty square in the generated PDF
		/// file.<br/>Normally this means that Excel is replacing fonts, for example Arial with MS Mincho, and to fix this error you should provide a 
		/// suitable list of Fallback fonts. Look at the PDF documentation in Exporting
		/// to PDF for more information ("Dealing with missing fonts and glyps" section).
		/// </summary>
		PdfGlyphNotInFont = 0x301,

		/// <summary>
		/// FlexCel is trying to render a character that is not in the font, but it is in the fallback font list. This character will be drawn with the
		/// fallback font.<br/>Normally this means that Excel is replacing fonts, for example Arial with MS Mincho. Look at the PDF documentation in Exporting
		/// to PDF for more information ("Dealing with missing fonts and glyps" section).
		/// </summary>
		PdfUsedFallbackFont = 0x302,

		/// <summary>
		/// The font you are trying to ue does not have a "bold" or "italic" variation, and we will use "faux" font created by making the normal font heavier
		/// or slanting it to the right. This normally results in lower quality fonts, and will not work with fonts embedded. We recomend that you use
		/// fonts that have italics and bold variations. Look at the PDF documentation in Exporting
		/// to PDF for more information ("Dealing with missing fonts and glyps" section).
		/// </summary>
		PdfFauxBoldOrItalics = 0x303,

		/// <summary>
		/// The "Fonts" folder in this machine has a file that could not be parsed by FlexCel and is probably corrupt. While this will not affect
		/// FlexCel in any way, you might want to look at the "Font" folder and remove this font, since it can make windows slower.
		/// </summary>
		PdfCorruptFontInFontFolder = 0x304,

		/// <summary>
		/// There was a sharing violation when trying to save an html file, and <see cref="FlexCel.Render.FlexCelHtmlExport.IgnoreSharingViolations"/> is true.
		/// Normally this just means two parallel threads trying to write the same file and can be safely ignored.
		/// </summary>
		HtmlSaveSharingViolation = 0x400,

        /// <summary>
        /// A malformed Url was detected in the xls file and was not exported. An example of this might be: "mailto:test@test@test".
        /// </summary>
        MalformedUrl = 0x500
	}

	#endregion

	#region FlexCelTrace
    /// <summary>
    /// This class reports al FlexCel non-fatal errors. Use it to diagnose when something is going wrong.
    /// </summary>
    public sealed class FlexCelTrace
    {
        private static bool FEnabled = true;

        private FlexCelTrace() {}

		internal static void Write(TFlexCelErrorInfo message)
		{
			if (OnError != null) OnError(message);
		}

		/// <summary>
		/// Set this to false if you want to prevent FlexCel from tracing non fatal errors. Note that if you
		/// don't have any event attached to this class the result will be the same as having Enabled = false.
		/// </summary>
        public static bool Enabled { get { return FEnabled; } set { FEnabled = value; } }

		internal static bool HasListeners {get{return Enabled && OnError != null;}}

		/// <summary>
		/// This event is called each time a non fatal error happens in FlexCel. Hook an event listener to it to be notified when this happens.
		/// </summary>
		public static event FlexCelErrorEventHandler OnError;
    }

	#endregion

	#region OnError event handler
	/// <summary>
	/// Delegate for ErrorInfo events.
	/// </summary>
	public delegate void FlexCelErrorEventHandler(TFlexCelErrorInfo e);

	#endregion

	#region Error information
	/// <summary>
	/// This class contains generic information about a non fatal error that happened in FlexCel.
	/// Children classes might contain more information specific to the error type.
	/// </summary>
	public abstract class TFlexCelErrorInfo
	{
		private FlexCelError FError;
		private string FMessage;

		/// <summary>
		/// Creates a new TFlexCelErrorInfo class.
		/// </summary>
		/// <param name="aError">Parameter indicating the type of error that happened.</param>
		/// <param name="aMessage">String detailing the error that happened.</param>
		protected TFlexCelErrorInfo(FlexCelError aError, string aMessage)
		{
			FError = aError;
			FMessage = aMessage;
		}

		/// <summary>
		/// Error type that this class is holding.
		/// </summary>
		public FlexCelError Error {get {return FError;}}

		/// <summary>
		/// Error message with detailed information on what happened.
		/// </summary>
		public string Message {get {return FMessage;}}
	}



	/// <summary>
	/// This class has information for a <see cref="FlexCelError.XlsTooManyPageBreaks"/> error. Look at <see cref="FlexCelError.XlsTooManyPageBreaks"/>
	/// for more information.
	/// </summary>
	public class TXlsTooManyPageBreaksError: TFlexCelErrorInfo
	{
		private string FFileName;

		/// <summary>
		/// Creates a new instance.
		/// </summary>
		/// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
		/// <param name="aFileName">See <see cref="FileName"/></param>
		public TXlsTooManyPageBreaksError(string aMessage, string aFileName): base(FlexCelError.XlsTooManyPageBreaks, aMessage)
		{
			FFileName = aFileName;
		}

		/// <summary>
		/// File with too many page breaks.
		/// </summary>
		public string FileName {get {return FFileName;}}

	}

    /// <summary>
    /// This class has information for a <see cref="FlexCelError.XlsxInvalidName"/> error. Look at <see cref="FlexCelError.XlsxInvalidName"/>
    /// for more information.
    /// </summary>
    public class TXlsxInvalidNameError : TFlexCelErrorInfo
    {
        private string FFileName;
        private string FName;
        private string FDefinition;

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        /// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
        /// <param name="aFileName">See <see cref="FileName"/></param>
        /// <param name="aName">See <see cref="Name"/></param>
        /// <param name="aDefinition">See <see cref="Definition"/></param>
        public TXlsxInvalidNameError(string aMessage, string aFileName, string aName, string aDefinition)
            : base(FlexCelError.XlsxInvalidName, aMessage)
        {
            FFileName = aFileName;
            FName = aName;
            FDefinition = aDefinition;
        }

        /// <summary>
        /// File with the invalid name.
        /// </summary>
        public string FileName { get { return FFileName; } }

        /// <summary>
        /// Name of the invalid named range.
        /// </summary>
        public string Name { get { return FName; } }

        /// <summary>
        /// Definition of the invalid named range.
        /// </summary>
        public string Definition { get { return FDefinition; } }


    }

    /// <summary>
    /// This class has information for a <see cref="FlexCelError.XlsxMissingPart"/> error. Look at <see cref="FlexCelError.XlsxMissingPart"/>
    /// for more information.
    /// </summary>
    public class TXlsxMissingPartError : TFlexCelErrorInfo
    {
        private string FFileName;
        private string FPartName;

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        /// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
        /// <param name="aFileName">See <see cref="FileName"/></param>
        /// <param name="aPartName">See <see cref="PartName"/></param>
        public TXlsxMissingPartError(string aMessage, string aFileName, string aPartName)
            : base(FlexCelError.XlsxMissingPart, aMessage)
        {
            FFileName = aFileName;
            FPartName = aPartName;
        }

        /// <summary>
        /// File with the missing part.
        /// </summary>
        public string FileName { get { return FFileName; } }

        /// <summary>
        /// Name of the missing part.
        /// </summary>
        public string PartName { get { return FPartName; } }

    }

	/// <summary>
	/// This class has information for a <see cref="FlexCelError.RenderMetafile"/> error. Look at <see cref="FlexCelError.RenderMetafile"/>
	/// for more information.
	/// </summary>
	public class TRenderMetafileError: TFlexCelErrorInfo
	{
		/// <summary>
		/// Creates a new instance.
		/// </summary>
        /// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
        public TRenderMetafileError(string aMessage) : base(FlexCelError.RenderMetafile, aMessage) { }
	}

	/// <summary>
	/// This class has information for a <see cref="FlexCelError.RenderCorruptImage"/> error. Look at <see cref="FlexCelError.RenderCorruptImage"/>
	/// for more information.
	/// </summary>
	public class TRenderCorruptImageError: TFlexCelErrorInfo
	{
		/// <summary>
		/// Creates a new instance.
		/// </summary>
        /// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
        public TRenderCorruptImageError(string aMessage) : base(FlexCelError.RenderCorruptImage, aMessage) { }
	}

	/// <summary>
	/// This class has information for a <see cref="FlexCelError.RenderErrorDrawingImage"/> error. Look at <see cref="FlexCelError.RenderErrorDrawingImage"/>
	/// for more information.
	/// </summary>
	public class TRenderErrorDrawingImageError: TFlexCelErrorInfo
	{
		/// <summary>
		/// Creates a new instance.
		/// </summary>
        /// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
		public TRenderErrorDrawingImageError(string aMessage): base(FlexCelError.RenderErrorDrawingImage, aMessage){}
	}

	/// <summary>
	/// This class has information for a <see cref="FlexCelError.PdfFontNotFound"/> error. Look at <see cref="FlexCelError.PdfFontNotFound"/>
	/// for more information.
	/// </summary>
	public class TPdfFontNotFoundError: TFlexCelErrorInfo
	{
		private readonly string FFontName;
		private readonly string FReplacementFontName;

		/// <summary>
		/// Creates a new instance.
		/// </summary>
        /// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
        /// <param name="aFontName">See <see cref="FontName"/></param>
        /// <param name="aReplacementFontName">See <see cref="ReplacementFontName"/></param>
		public TPdfFontNotFoundError(string aMessage, string aFontName, string aReplacementFontName): base(FlexCelError.PdfFontNotFound, aMessage)
		{
			FFontName = aFontName;
			FReplacementFontName = aReplacementFontName;
		}

		/// <summary>
		/// Font that was not found.
		/// </summary>
		public string FontName {get {return FFontName;}}

		/// <summary>
		/// Font that was used to replace <see cref="FontName"/>.
		/// </summary>
		public string ReplacementFontName {get {return FReplacementFontName;}}

	}

	/// <summary>
	/// This class has information for a <see cref="FlexCelError.PdfUsedFallbackFont"/> error. Look at <see cref="FlexCelError.PdfUsedFallbackFont"/>
	/// for more information.
	/// </summary>
	public class TPdfUsedFallbackFontError: TFlexCelErrorInfo
	{
		private readonly string FOriginalFontName;
		private readonly string FSubstitutedFontName;

		/// <summary>
		/// Creates a new instance.
		/// </summary>
        /// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
        /// <param name="aOriginalFontName">See <see cref="OriginalFontName"/></param>
		/// <param name="aSubstitutedFontName">See <see cref="SubstitutedFontName"/></param>
		public TPdfUsedFallbackFontError(string aMessage, string aOriginalFontName, string aSubstitutedFontName): base(FlexCelError.PdfUsedFallbackFont, aMessage)
		{
			FOriginalFontName = aOriginalFontName;
			FSubstitutedFontName = aSubstitutedFontName;
		}

		/// <summary>
		/// Font that should be used, but that doesn't contain the needed characters.
		/// </summary>
		public string OriginalFontName {get {return FOriginalFontName;}}

		/// <summary>
		/// Fallback font that substituted <see cref="OriginalFontName"/>.
		/// </summary>
		public string SubstitutedFontName {get {return FSubstitutedFontName;}}

	}

	/// <summary>
	/// This class has information for a <see cref="FlexCelError.PdfGlyphNotInFont"/> error. Look at <see cref="FlexCelError.PdfGlyphNotInFont"/>
	/// for more information.
	/// </summary>
	public class TPdfGlyphNotInFontError: TFlexCelErrorInfo
	{
		private readonly string FFontName;
		private readonly long FMissingChar;

		/// <summary>
		/// Creates a new instance.
		/// </summary>
		/// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
		/// <param name="aFontName">See <see cref="FontName"/></param>
		/// <param name="aMissingChar">See <see cref="MissingChar"/></param>
		public TPdfGlyphNotInFontError(string aMessage, string aFontName, long aMissingChar): base(FlexCelError.PdfGlyphNotInFont, aMessage)
		{
			FFontName = aFontName;
			FMissingChar = aMissingChar;
		}

		/// <summary>
		/// Character missing in the font.
		/// </summary>
		public long MissingChar {get {return FMissingChar;}}

		/// <summary>
		/// Font that doesn't contain the character.
		/// </summary>
		public string FontName {get {return FFontName;}}

	}

	/// <summary>
	/// This class has information for a <see cref="FlexCelError.PdfFauxBoldOrItalics"/> error. Look at <see cref="FlexCelError.PdfFauxBoldOrItalics"/>
	/// for more information.
	/// </summary>
	public class TPdfFauxBoldOrItalicsError: TFlexCelErrorInfo
	{
        private string FFontName;
		private FontStyle FStyle;

		/// <summary>
		/// Creates a new instance.
		/// </summary>
        /// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
        /// <param name="aFontName">See <see cref="FontName"/></param>
		/// <param name="aStyle">See <see cref="Style"/></param>
		public TPdfFauxBoldOrItalicsError(string aMessage, string aFontName, FontStyle aStyle): base(FlexCelError.PdfFauxBoldOrItalics, aMessage)
		{
			FFontName = aFontName;
			FStyle = aStyle;
		}

		/// <summary>
		/// Name of the font that doesn't contain Italics or bold definition.
		/// </summary>
		public string FontName {get {return FFontName;} set{FFontName = value;}}

		/// <summary>
		/// Style missing from the font.
		/// </summary>
		public FontStyle Style {get {return FStyle;} set{FStyle = value;}}

	}

	/// <summary>
	/// This class has information for a <see cref="FlexCelError.PdfCorruptFontInFontFolder"/> error. Look at <see cref="FlexCelError.PdfCorruptFontInFontFolder"/>
	/// for more information.
    /// </summary>
    public class TPdfCorruptFontInFontFolderError : TFlexCelErrorInfo
	{
		private string FFileName;

		/// <summary>
		/// Creates a new instance.
		/// </summary>
        /// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
        /// <param name="aFileName">See <see cref="FileName"/></param>
		public TPdfCorruptFontInFontFolderError(string aMessage, string aFileName): base(FlexCelError.PdfCorruptFontInFontFolder, aMessage)
		{
			FFileName = aFileName;
		}

		/// <summary>
		/// Font file that FlexCel couldn't parse.
		/// </summary>
		public string FileName {get {return FFileName;}}

	}

	/// <summary>
	/// This class has information for a <see cref="FlexCelError.HtmlSaveSharingViolation"/> error. Look at <see cref="FlexCelError.HtmlSaveSharingViolation"/>
	/// for more information.
    /// </summary>
    public class THtmlSaveSharingViolationError : TFlexCelErrorInfo
	{
		private string FFileName;

		/// <summary>
		/// Creates a new instance.
		/// </summary>
        /// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
        /// <param name="aFileName">See <see cref="FileName"/></param>
		public THtmlSaveSharingViolationError(string aMessage, string aFileName): base(FlexCelError.HtmlSaveSharingViolation, aMessage)
		{
			FFileName = aFileName;
		}

		/// <summary>
		/// File with the sharing violation.
		/// </summary>
		public string FileName {get {return FFileName;}}

	}


    /// <summary>
    /// This class has information for a <see cref="FlexCelError.MalformedUrl"/> error. Look at <see cref="FlexCelError.MalformedUrl"/>
    /// for more information.
    /// </summary>
    public class TMalformedUrlError : TFlexCelErrorInfo
    {
        private string FUrl;

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        /// <param name="aMessage">See <see cref="TFlexCelErrorInfo.Message"/></param>
        /// <param name="aUrl">See <see cref="Url"/></param>
        public TMalformedUrlError(string aMessage, string aUrl) : base(FlexCelError.MalformedUrl, aMessage) 
        {
            FUrl = aUrl;
        }

        /// <summary>
        /// Malformed url.
        /// </summary>
        public string Url { get { return FUrl; } }

    }

	#endregion

}
