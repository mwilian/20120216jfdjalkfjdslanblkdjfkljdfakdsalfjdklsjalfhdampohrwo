#region Using directives

using System;
using System.Text;
using System.Globalization;
using System.Diagnostics;
using System.Collections.Generic;

#if (WPF)
using RectangleF = System.Windows.Rect;
using System.Windows.Media;
#else
using System.Drawing;
#endif


#endregion

namespace FlexCel.Pdf
{
    #region Font options
    /// <summary>
    /// The way fonts will be embedded on the resulting pdf file.
    /// </summary>
    public enum TFontEmbed
    {
        /// <summary>
        /// No font will be embedded. The result file will be smaller, but the file might not look fine on a
        /// computer without the font installed. It is recomended that you embed the fonts.
        /// </summary>
        None,

        /// <summary>
        /// All fonts will be embedded. The file will be larger than when not embedding fonts, but it will print on any computer. 
        /// Note that you can control which fonts to embed wnd which not with the OnFontEmbed event.
        /// </summary>
        Embed,

        /// <summary>
        /// This is a compromise between embedding all fonts and not embedding them. It will only embed fonts with symbols, and leave normal fonts
        /// not embedded.
        /// </summary>
        OnlySymbolFonts
    }

	/// <summary>
	/// Determines if full fonts will be embedded in the generated pdf files, or only the characters being used.
	/// </summary>
	public enum TFontSubset
	{
		/// <summary>
		/// All characters of the font will be embedded into the file. This setting creates bigger files, but they can be edited after generated.
		/// </summary>
		DontSubset,

		/// <summary>
		/// Only characters actually used by the document will be embedded into the file. This will create smaller files than 
		/// embedding the full font, but it will be difficult to edit the document once it has been created. 
		/// </summary>
		Subset
	}


    /// <summary>
    /// How fonts will be replaced on the generated PDF file.
    /// </summary>
    public enum TFontMapping
    {
        /// <summary>
        /// Arial will be replaced with Helvetica, Times new roman with Times and True type Courier with
        /// PS1 Courier. All other fonts will remain unchanged.
        /// </summary>
        ReplaceStandardFonts,

        /// <summary>
        /// Serif fonts will be mapped to Times, MonoSpace to Courier, Sans-Serif to Helvetica and Symbol fonts to Symbol. Using this option
        /// you can get the smallest file sizes and 100% portability, but the resulting file will only use those fonts. Use it with care, specially if you use symbol fonts.
        /// </summary>
        ReplaceAllFonts,

        /// <summary>
        /// All actual fonts will be used. If you use this option and do not embed fonts, the fonts will look bad on computers
        /// without them installed. If you embed fonts, files will be larger.
        /// </summary>
        DontReplaceFonts
    }

    #endregion

	internal class TTracedFonts: Dictionary<string, string> {}

    internal sealed class PdfConv
    {
        private PdfConv(){}

        public static string CoordsToString(double d)
        {
			Debug.Assert(!double.IsNaN(d));
			if ( d > 4000000) d = 4000000;
			if ( d < -4000000) d = -4000000;
            return d.ToString("0.####", CultureInfo.InvariantCulture);
        }

		public static string DoubleToString(double d)
		{
			Debug.Assert(!double.IsNaN(d));
			return d.ToString("0.####", CultureInfo.InvariantCulture);
		}

		public static string LongToString(long d)
		{
			return d.ToString(CultureInfo.InvariantCulture);
		}

        public static string ToString(Color c)
        {
            return CoordsToString(c.R / 255f) + " " + CoordsToString(c.G / 255f)+ " " + CoordsToString(c.B / 255f);
        }

        public static string ToRectangleXY(RectangleF Coords, bool AddBrackets)
        {
            StringBuilder Result = new StringBuilder();
            if (AddBrackets) Result.Append(TPdfTokens.GetString(TPdfToken.OpenArray));
            Result.Append(CoordsToString(Coords.Left)); Result.Append(" ");
            Result.Append(CoordsToString(Coords.Top)); Result.Append(" ");
            Result.Append(CoordsToString(Coords.Right)); Result.Append(" ");
            Result.Append(CoordsToString(Coords.Bottom));  
            if (AddBrackets) Result.Append(TPdfTokens.GetString(TPdfToken.CloseArray));
            return Result.ToString();
        }

        public static string ToRectangleWH(RectangleF Coords, bool AddBrackets)
        {
            StringBuilder Result = new StringBuilder();
            if (AddBrackets) Result.Append(TPdfTokens.GetString(TPdfToken.OpenArray));
            Result.Append(CoordsToString(Coords.Left)); Result.Append(" ");
            Result.Append(CoordsToString(Coords.Top)); Result.Append(" ");
            Result.Append(CoordsToString(Coords.Width)); Result.Append(" ");
            Result.Append(CoordsToString(Coords.Height));
            if (AddBrackets) Result.Append(TPdfTokens.GetString(TPdfToken.CloseArray));
            return Result.ToString();
        }

		public static string ToString(double[] Source, bool AddBrackets)
		{
			StringBuilder Result = new StringBuilder();
			if (AddBrackets) Result.Append(TPdfTokens.GetString(TPdfToken.OpenArray));
			for (int i = 0; i < Source.Length; i++)
			{
				Result.Append(PdfConv.CoordsToString(Source[i])); 
				if (i < Source.Length - 1) Result.Append(" ");
			}
			if (AddBrackets) Result.Append(TPdfTokens.GetString(TPdfToken.CloseArray));
			return Result.ToString();
		}

		public static string ToString(float[] Source, bool AddBrackets)
		{
			StringBuilder Result = new StringBuilder();
			if (AddBrackets) Result.Append(TPdfTokens.GetString(TPdfToken.OpenArray));
			for (int i = 0; i < Source.Length; i++)
			{
				Result.Append(PdfConv.CoordsToString(Source[i])); 
				if (i < Source.Length - 1) Result.Append(" ");
			}
			if (AddBrackets) Result.Append(TPdfTokens.GetString(TPdfToken.CloseArray));
			return Result.ToString();
		}

        public static string ToString(int[] Source, bool AddBrackets)
        {
            StringBuilder Result = new StringBuilder();
            if (AddBrackets) Result.Append(TPdfTokens.GetString(TPdfToken.OpenArray));
            for (int i = 0; i < Source.Length; i++)
            {
                Result.Append(PdfConv.LongToString(Source[i]));
                if (i < Source.Length - 1) Result.Append(" ");
            }
            if (AddBrackets) Result.Append(TPdfTokens.GetString(TPdfToken.CloseArray));
            return Result.ToString();
        }

        public static string ToHexString(byte[] b, bool AddDelims)
        {
            StringBuilder Result = new StringBuilder(b.Length * 2 + 4); //it really is +2, but just to be safe
            if (AddDelims) Result.Append("<");
            for (int i = 0; i < b.Length; i++)
            {
                Result.AppendFormat("{0:X2}", b[i]);
            }
            if (AddDelims) Result.Append(">");

            return Result.ToString();
        }

    }

    /// <summary>
    /// Encapsulates the document properties for the PDF file.
    /// </summary>
    public class TPdfProperties
    {
        private string FTitle;
        private string FAuthor;
        private string FSubject;
        private string FKeywords;
        private string FCreator;


        /// <summary>
        /// Creates a new instance of the class. All properties are set to null.
        /// </summary>
        public TPdfProperties()
        {
        }

        /// <summary>
        /// Creates a new instance of the class with given properties.
        /// </summary>
        /// <param name="aTitle">Document title.</param>
        /// <param name="aAuthor">Document author.</param>
        /// <param name="aSubject">Document subject.</param>
        /// <param name="aKeywords">Keywords to search on the document.</param>
        /// <param name="aCreator">Application that created the document.</param>
        public TPdfProperties(string aTitle,
                              string aAuthor,
                              string aSubject,
                              string aKeywords,
                              string aCreator)
        {
            FTitle = aTitle;
            FAuthor = aAuthor;
            FSubject = aSubject;
            FKeywords = aKeywords;
            FCreator = aCreator;
        }

        /// <summary>
        /// Document title.
        /// </summary>
        public string Title { get { return FTitle; } set { FTitle = value; } }

        /// <summary>
        /// Document author.
        /// </summary>
        public string Author { get { return FAuthor; } set { FAuthor = value; } }

        /// <summary>
        /// Document subject.
        /// </summary>
        public string Subject { get { return FSubject; } set { FSubject = value; } }

        /// <summary>
        /// Keywords to search on the document.
        /// </summary>
        public string Keywords { get { return FKeywords; } set { FKeywords = value; } }

        /// <summary>
        /// Application that created the document.
        /// </summary>
        public string Creator { get { return FCreator; } set { FCreator = value; } }

    }

	/// <summary>
	/// Icon for a pdf comment
	/// </summary>
	public enum TPdfCommentIcon
	{
		/// <summary>A callout icon..</summary>
		Comment,

		/// <summary>A ? inside a circle icon.</summary>
		Help,

		/// <summary>A triangle icon.</summary>
		Insert,

		/// <summary>A key icon.</summary>
		Key,

		/// <summary>A small triangle icon.</summary>
		NewParagraph,

		/// <summary>A sheet of paper icon.</summary>
		Note,

		/// <summary>A "Pi" icon.</summary>
		Paragraph
	}

	/// <summary>
	/// Different types of comments.
	/// </summary>
	public enum TPdfCommentType
	{
		/// <summary>
		/// An icon that will show the comment.
		/// </summary>
		Text,

		/// <summary>
		/// A rectangle including the comment.
		/// </summary>
		Square,

		/// <summary>
		/// A Circle or ellipsis including the comment.
		/// </summary>
		Circle
	}

	/// <summary>
	/// Properties for a PDF comment.
	/// </summary>
	public class TPdfCommentProperties
	{
		private TPdfCommentType FCommentType;

		private TPdfCommentIcon FIcon;
		private float FOpacity;
		private Color FBackgroundColor;
		private Color FLineColor;

		/// <summary>
		/// Creates a new instance of a comment object.
		/// </summary>
		/// <param name="aCommentType"></param>
		/// <param name="aIcon"></param>
		/// <param name="aOpacity"></param>
		/// <param name="aBackgroundColor"></param>
		/// <param name="aLineColor"></param>
		public TPdfCommentProperties(TPdfCommentType aCommentType, TPdfCommentIcon aIcon, float aOpacity, Color aBackgroundColor, Color aLineColor)
		{
			FCommentType	 = aCommentType;				   
			FIcon 			 = aIcon;
			FOpacity     	 = aOpacity;
			FBackgroundColor = aBackgroundColor;
			FLineColor		 = aLineColor;
		}

		/// <summary>
		/// Creates a new TPdfCommentProperties instance based on the data form another instance.
		/// </summary>
		/// <param name="aProps"></param>
		public TPdfCommentProperties(TPdfCommentProperties aProps)
		{
			FCommentType	 = aProps.CommentType;				   
			FIcon 			 = aProps.Icon;
			FOpacity     	 = aProps.Opacity;
			FBackgroundColor = aProps.BackgroundColor;
			FLineColor		 = aProps.LineColor;
		}

		/// <summary>
		/// Type of comment.
		/// </summary>
		public TPdfCommentType CommentType {get{return FCommentType;} set{FCommentType=value;}}

		/// <summary>
		/// Icon for the comment. Only visible if <see cref="CommentType"/> is Text
		/// </summary>
		public TPdfCommentIcon Icon {get{return FIcon;} set{FIcon=value;}}

		/// <summary>
		/// A value between 0 and 1 specifying the opacity of the note.
		/// </summary>
		public float Opacity {get{return FOpacity;} set{FOpacity=value;}}

		/// <summary>
		/// Background color for the comment. Only visible if <see cref="CommentType"/> is NOT Text
		/// </summary>
		public Color BackgroundColor {get{return FBackgroundColor;} set{FBackgroundColor=value;}}

		/// <summary>
		/// Line color for the comment. Only visible if <see cref="CommentType"/> is NOT Text
		/// </summary>
		public Color LineColor {get{return FLineColor;} set{FLineColor=value;}}

	}

	/// <summary>
	/// Viewer settings when the document is opened for the first time.
	/// </summary>
	public enum TPageLayout
	{
        /// <summary>
        /// Keep the layout as defined by the user.
        /// </summary>
		None,

        /// <summary>
        /// Show the Outlines pane when opening the document.
        /// </summary>
		Outlines,

        /// <summary>
        /// Show the thumbs pane when opening the document.
        /// </summary>
		Thumbs,

        /// <summary>
        /// Open the document in FullScreen mode.
        /// </summary>
		FullScreen
	}

}
