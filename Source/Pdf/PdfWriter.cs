#region Using directives

using System;
using System.Text;
using System.IO;
using System.Globalization;
using FlexCel.Core;
using System.Collections.Generic;

#if (WPF)
using RectangleF = System.Windows.Rect;
using SizeF = System.Windows.Size;
using real = System.Double;
using System.Windows.Media;
#else
using real = System.Single;
using Colors = System.Drawing.Color;
using DashStyles = System.Drawing.Drawing2D.DashStyle;
using System.Drawing;
using System.Drawing.Drawing2D;
#endif



#endregion

namespace FlexCel.Pdf
{
    /// <summary>
    /// A simple class for creating PDF files. It will not hold contents into memory, so it can be used with little memory.
    /// </summary>
    /// <remarks>
    /// This class is not intended for providing a complete API for writing PDFs, only
    /// what is necessary to create them from xls files.
    /// Even when this class could be used standalone, on most cases <see cref="FlexCel.Render.FlexCelPdfExport"/> should be used.
    /// </remarks>
	public class PdfWriter
	{
		#region Private variables
		private bool FCompress;
		private TPdfStream DataStream;
		private TPdfSignature FSignature;

        internal const bool FTesting = false; //Only for testing, should be true on normal work.

		private const real InchesToUnits=72F/100F;

		TBrushStyle LastBrushStyle;
		Color LastBrushColor;
		HatchStyle LastHatchStyle;
		Color LastPenColor;
		real LastPenWidth;
		double[] DrawingMatrix;
		DashStyle LastPenStyle;
		Stack<TGState> GraphicsState;

		string LastFont;
		real LastTextY;
		real LastTextX;

		TPaperDimensions FPageSize;

		TPdfProperties FProperties;

		TFontEmbed FFontEmbed;
		TFontSubset FFontSubset;
		TFontMapping FFontMapping;
		bool FKerning;
		string FFallbackFonts;

		TBodySection Body;
		TBodySection TempBody;
		TXRefSection XRef;

		real FScale;
		bool FYAxisGrowsDown;
		bool FAddFontDescent;

		TFontEvents FontEvents;

		TTracedFonts TracedFonts;

		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new PDF file instance.
		/// </summary>
		public PdfWriter()
		{
			FCompress = true;
			FProperties = new TPdfProperties();
			GraphicsState = new Stack<TGState>();
			FPageSize = DefaultPaperSize;
			FYAxisGrowsDown = false;
			FScale = 1;
			FFontEmbed = TFontEmbed.None;
			FFontSubset = TFontSubset.Subset;
			FFontMapping = TFontMapping.ReplaceStandardFonts;

			FontEvents = new TFontEvents(this, null, null, null);

			FFallbackFonts = "Arial Unicode MS";
		}

		internal static readonly TPaperDimensions DefaultPaperSize= new TPaperDimensions("A4",827,1169);  //STATIC*

		#endregion

		#region Properties
		/// <summary>
		/// Set it to true to compress the text on the generated pdf file. 
		/// </summary>
		/// <value></value>
		public bool Compress
		{get{return FCompress;} set{FCompress = value;}}

		/// <summary>
		/// Properties of the PDF file.
		/// </summary>
		/// <value></value>
		public TPdfProperties Properties
		{ get { return FProperties; } set { if (value == null) FProperties = new TPdfProperties(); else FProperties = value; } }

		/// <summary>
		/// Page size of the active page. You can change it *before* calling NewPage() and it will change for the new sheets.
		/// Note that once NewPage() (or BeginDoc() for the first page) is called, the page size will remain constant for that page.
		/// This property must be changed before.
		/// </summary>
		public TPaperDimensions PageSize{get{return FPageSize;} set{if (value==null) FPageSize= DefaultPaperSize; else FPageSize=value;}}
        
		/// <summary>
		/// When true, the y axis origin corresponds to the upper corner of a sheet and bigger Y coordinates will 
		/// move down on the paper. This is the standard GDI+ behavior.
		/// When false, the Y origin is at the bottom and it grows up into the page. This is the standard PDF behavior.
		/// </summary>
		public bool YAxisGrowsDown {get {return FYAxisGrowsDown;} set {FYAxisGrowsDown = value;}}
        
		/// <summary>
		/// When false, (the default) text base will be at the y coordinate. For example, DrawString(..., y=100,...)
		/// will draw a string with its base at 100. Font descent (for example the lower part of a "p") will be below
		/// 100, and the ascent (the upper part) will be above. This is the standard PDF behavior.
		/// When true, all text will be drawn above the y coordinate. (both ascent and descent).
		/// This is the standard GDI+ behavior, when StringFormat.LineAlignment=StringAlignment.Far.
		/// </summary>
		public bool AddFontDescent {get {return FAddFontDescent;} set {FAddFontDescent = value;}}

		/// <summary>
		/// A scale factor to change X and Y coordinates. When Scale=1, the using is the point (1/72 of an inch).
		/// Font size is not affected by scale, it is always in points.
		/// </summary>
		public real Scale {get {return FScale;} set {FScale=value;}}

		/// <summary>
		/// Determines what fonts will be embedded on the generated pdf. 
		/// Note that when using UNICODE fonts WILL BE EMBEDDED no matter the value of this property.
		/// </summary>
		public TFontEmbed FontEmbed {get {return FFontEmbed;} set {FFontEmbed=value;}}

		/// <summary>
		/// When <see cref="FontEmbed"/> is set to embed the fonts, this setting determines if the full font will be embedded,
		/// or only the characters used in the document. If the full font is embedded the resulting file will be larger, but 
		/// it will be possible to edit it with a third party tool once it has been generated.
		/// </summary>
		public TFontSubset FontSubset {get {return FFontSubset;} set {FFontSubset=value;}}

		/// <summary>
		/// Determines how fonts will be replaced on the generated pdf. Pdf comes with 4 standard font families,
		/// Serif, Sans-Serif, Monospace and Symbol. You can use for example the standard Helvetica instead of Arial and do not worry about embedding the font.
		/// </summary>
		public TFontMapping FontMapping {get {return FFontMapping;} set {FFontMapping=value;}}

		/// <summary>
		/// By default, pdf does not do any kerning with the fonts. This is, on the string "AVANT", it won't
		/// compensate the spaces between "A" and "V". (they should be smaller) 
		/// If you turn this property on, FlexCel will calculate the kerning and add it to the generated file.
		/// The result file will be a little bigger because of the kerning info on all strings, but it will also
		/// look a little better.
		/// </summary>
		public bool Kerning {get {return FKerning;} set {FKerning=value;}}

		/// <summary>
		/// A semicolon (;) separated list of font names to try when a character is not found in the used font.<br/>
		/// When a character is not found in a font, it will display as an empty square by default. By setting this
		/// property, FlexCel will try to find a font that supports this character in this list, and if found, use that font
		/// to render the character.
		/// </summary>
		/// <example>
		/// You might set this property to "MS MINCHO;Arial Unicode MS". If a cell with font "Arial" has a character that
		/// is not in the "Arial" font, FlexCel will try to find the character first in MS Mincho, and if not found, in Arial Unicode.<br></br>
		/// If it can find it in any of the fallback fonts, it will use that font in the pdf file.
		/// </example>
		public string FallbackFonts {get {return FFallbackFonts;} set {FFallbackFonts=value;}}

		/// <summary>
		/// Use this event if you want to provide your own font information for embedding. 
		/// Note that if you don't assign this event, the default method will be used, and this 
		/// will try to find the font on the Fonts folder. To change the font folder, use <see cref="GetFontFolder"/> event
		/// </summary>
		public GetFontDataEventHandler GetFontData {get {return FontEvents.OnGetFontData;} set { FontEvents.OnGetFontData = value;}}

		/// <summary>
		/// Use this event if you want to provide your own font information for embedding. 
		/// Normally FlexCel will search for fonts on [System]\Fonts folder. If your fonts are in 
		/// other location, you can tell FlexCel where they are here. If you prefer just to give FlexCel
		/// the full data on the font, you can use <see cref="GetFontData"/> event instead.
		/// </summary>
		public GetFontFolderEventHandler GetFontFolder {get {return FontEvents.OnGetFontFolder;} set { FontEvents.OnGetFontFolder = value;}}

		/// <summary>
		/// Use this event if you want to manually specify which fonts to embed into the pdf document.
		/// </summary>
		public FontEmbedEventHandler OnFontEmbed {get {return FontEvents.OnFontEmbed;} set { FontEvents.OnFontEmbed = value;}}

        
		#endregion

		#region Document Managment

		/// <summary>
		/// Call this method before starting the output.
		/// It will initialize a new page. After this you can call <see cref="DrawString(string, Font, Brush, real, real)"/>, <see cref="NewPage"/>, etc.
		/// Always end the document with a call to <see cref="EndDoc"/> and then remember to close the stream.
		/// </summary>
		public void BeginDoc(Stream aDataStream)
		{
			DataStream = new TPdfStream(aDataStream, FSignature);
			THeaderSection.SaveToStream(DataStream);
			Body = new TBodySection(Compress);
			TempBody = null; //just to release resources.
			XRef = new TXRefSection();
			LastBrushColor = Colors.Black;
			LastHatchStyle = (HatchStyle)(-1);
			LastBrushStyle = TBrushStyle.Solid;
			LastPenColor = Colors.Black;
			LastPenWidth = -1;
			LastPenStyle = DashStyles.Solid;
			DrawingMatrix = new double[]{1,0,0,1,0,0};
			TracedFonts = new TTracedFonts();

			Body.BeginSave(DataStream, XRef, PageSize, FSignature, FallbackFonts, FontEvents);
		}

		/// <summary>
		/// Always call this method to write the final part of a pdf file.
		/// </summary>
		public void EndDoc()
		{
			if (Body == null) return; //Enddoc has already been called, or begindoc was never called.

			TracedFonts = null;
			Body.EndSave(DataStream, XRef, Properties, FSignature);
			XRef.SaveToStream(DataStream);
			TTrailerSection.SaveToStream(DataStream, XRef, Body.CatalogId, Body.InfoId);
			Body.FinishSign(DataStream); //Last thing before the end.
			if (DataStream != null) DataStream.Dispose();
			DataStream = null;
			Body = null;
		}

		/// <summary>
		/// Closes the active page and creates a new one. All following commands will go to the new page.
		/// </summary>
		public void NewPage()
		{
			LastBrushColor = Colors.Black;
			LastHatchStyle = (HatchStyle)(-1);
			LastBrushStyle = TBrushStyle.Solid;

			LastPenColor = Colors.Black;
			LastPenWidth = -1;
			DrawingMatrix = new double[]{1,0,0,1,0,0};
			LastPenStyle = DashStyles.Solid;
			Body.NewPage(DataStream, XRef, PageSize);
		}
		#endregion

		#region Conversion
		private string tx(real x)
		{
			return PdfConv.CoordsToString(_tx(x));
		}

		private real _tx(real x)
		{
			return x*Scale;
		}
        
		/// <summary>
		/// This implementation is similar to tx, but concept is different. This is used to change
		/// a unit like width or height. If we introduced a XAxisGrowsRight, it would be different.
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		private string t(real s)
		{
			return PdfConv.CoordsToString(_t(s));
		}

		private real _t(real s)
		{
			return s*Scale;
		}
        
		private string ty(real y)
		{
			return PdfConv.CoordsToString(_ty(y));
		}

		private real _ty(real y)
		{
			if (YAxisGrowsDown)
				return (PageSize.Height*InchesToUnits*CurrentYScale - y*Scale);
			else
				return (y)*Scale;
		}
       
		private string thy(real y, real height)
		{
			return PdfConv.CoordsToString(_thy(y, height));
		}

		private real _thy(real y, real height)
		{
			if (YAxisGrowsDown)
				return _ty(y+height);
			else
				return _ty(y);
		}

		private string tnegy(real y)
		{
			return PdfConv.CoordsToString(_tnegy(y));
		}

		private real _tnegy(real y)
		{
			if (YAxisGrowsDown)
				return -(PageSize.Height*InchesToUnits*CurrentYScale - y*Scale);
			else
				return -(y)*Scale;
		}

		private real ti(real f)
		{
			return f/Scale;
		}

		private real tinvy(real y)
		{
			if (YAxisGrowsDown)
			{
				return (PageSize.Height*InchesToUnits*CurrentYScale - y) / Scale;
			}
			else
				return y/Scale;
		}

		#endregion

		#region Text

		/// <summary>
		/// Writes a string to the current page.
		/// </summary>
		/// <param name="x">X coord. (default from bottom left)</param>
		/// <param name="aFont">Font to draw the text.</param>
		/// <param name="aBrush">Brush used for Color.</param>
		/// <param name="y">Y coord. (default from bottom left. Might change with <see cref="YAxisGrowsDown"/> value)</param>
		/// <param name="s">String to write.</param>
		public void DrawString(string s, Font aFont, Brush aBrush, real x, real y)
		{
			DrawString(s, aFont, null, aBrush, x, y);
		}

		/// <summary>
		/// Writes a string to the current page.
		/// </summary>
		/// <param name="x">X coord. (default from bottom left)</param>
		/// <param name="aFont">Font to draw the text.</param>
		/// <param name="aPen">Pen to draw the text outline. If null, no outline will be drawn.</param>
		/// <param name="aBrush">Brush used for Color. If null, only the outline will be drawn.</param>
		/// <param name="y">Y coord. (default from bottom left. Might change with <see cref="YAxisGrowsDown"/> value)</param>
		/// <param name="s">String to write.</param>
		public void DrawString(string s, Font aFont, Pen aPen, Brush aBrush, real x, real y)
		{
			if (s == null || s.Length==0) return;
			SelectPen(aPen);

			bool SavedState = false;
			TPdfFont PdfFont = null;

			//Transparency (CommandGs) resets all parameters after text is drawn, this seems like a bug?
			//In this case, underline and strikeouts will render wrong, so we need to set this again.
			bool NeedstoRestoreState = (aFont.Underline || aFont.Strikeout) && BrushIsTransparent(aBrush);
			if (NeedstoRestoreState) SaveState(); 
			try
			{
 			    SelectBrush(aBrush, 0, out SavedState); //We are not doing 2 passes for hatches here.
				try
				{
					if (DataStream.PendingEndText) //This means pen and brush and anything else did not change, since the last command was "ET". we will concatenate this.
					{
						DataStream.PendingEndText = false; 
					}
					else
					{
						TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandBeginText));
						LastFont = null;
						LastTextY = 0;
						LastTextX = 0;
					}

					bool RestoreTextRendering = aPen != null;
					if (RestoreTextRendering)
					{
						if (aBrush != null) 
							TPdfBaseRecord.Write(DataStream, "2 ");
						else
							TPdfBaseRecord.Write(DataStream, "1 ");

						TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandTextRendering));
					}
            
					PdfFont = Body.SelectFont(DataStream, FFontMapping, aFont, s, FFontEmbed, FFontSubset, FKerning, ref LastFont);
            
					if (FAddFontDescent) y = y-FontDescent(PdfFont, aFont);

					real TextY = _ty(y);
					real TextX = _tx(x);

					TPdfBaseRecord.WriteLine(DataStream, PdfConv.CoordsToString(TextX - LastTextX)+" "+
						PdfConv.CoordsToString(TextY - LastTextY) + " " + TPdfTokens.GetString(TPdfToken.CommandTextMove));
            
					LastTextY = TextY;
					LastTextX = TextX;

					TKernedString[] ks = null;
					if (FKerning) 
						ks = PdfFont.KernString(s);
					if (ks == null || ks.Length <=1)
					{
						TPdfStringRecord.WriteStringInStream(DataStream, s, aFont.SizeInPoints, PdfFont, ref LastFont, TPdfTokens.GetString(TPdfToken.CommandTextWrite), String.Empty, String.Empty, TracedFonts);
					}
					else
					{
						TPdfBaseRecord.Write(DataStream, TPdfTokens.GetString(TPdfToken.OpenArray));
						for (int i=0; i<ks.Length; i++)
						{
							if (ks[i].Kern!=0) TPdfBaseRecord.Write(DataStream, Convert.ToString(-ks[i].Kern, CultureInfo.InvariantCulture));
							TPdfStringRecord.WriteStringInStream(DataStream, ks[i].Text, aFont.SizeInPoints, PdfFont, ref LastFont,
								String.Empty, TPdfTokens.GetString(TPdfToken.OpenArray), TPdfTokens.GetString(TPdfToken.CloseArray) + TPdfTokens.GetString(TPdfToken.CommandTextKerningWrite), TracedFonts);
						}
						TPdfBaseRecord.Write(DataStream, TPdfTokens.GetString(TPdfToken.CloseArray));
						TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandTextKerningWrite));
					}
            

					if (RestoreTextRendering)
					{
						TPdfBaseRecord.Write(DataStream, "0 ");
						TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandTextRendering));
					}


					//TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandEndText));
					DataStream.PendingEndText = true;
				}
				finally
				{
					if (NeedstoRestoreState) RestoreState();
				}
				if (aFont.Underline)
				{
					UnderLine(s, PdfFont, aFont.SizeInPoints, aBrush, x, y, false);
				}

				if (aFont.Strikeout)
				{
					UnderLine(s, PdfFont, aFont.SizeInPoints, aBrush, x, y, true);
				}
			}
			finally
			{
				if (SavedState) RestoreState();
			}
		}

		/// <summary>
		/// Underline has to be done by writing a line below the text. There is no PDF command for underlined text.
		/// </summary>
		/// <param name="s"></param>
		/// <param name="aPdfFont"></param>
		/// <param name="SizeInPoints"></param>
		/// <param name="aBrush"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		/// <param name="Strikeout"></param>
		private void UnderLine(string s, TPdfFont aPdfFont, real SizeInPoints, Brush aBrush, real x, real y, bool Strikeout)
		{
			if (aPdfFont == null) return;

			real sz = aPdfFont.MeasureString(s)* SizeInPoints * 0.001F;

			real Hang = (real)(SizeInPoints / 8F);
			if (aPdfFont.UnderlinePosition()!=0) Hang = -aPdfFont.UnderlinePosition()* SizeInPoints * 0.001F;

			if (Strikeout) Hang = -(real)(-Math.Abs(aPdfFont.Descent())+Math.Abs(aPdfFont.Ascent()))* SizeInPoints * 0.001F / 2F;

			if (YAxisGrowsDown) Hang=-Hang;
			Hang/=Scale;

			real LineWidth = (real)(SizeInPoints / 14.0 + 0.1)/Scale;		
	
			//We could eliminate the pen and use always brushes. But we keep it to keep the format compatible with older versions.
			SolidBrush SBrush = aBrush as SolidBrush;
			if (SBrush != null)
			{
				using (Pen p = new Pen(SBrush.Color, (real)(SizeInPoints / 14.0 + 0.1)/Scale))
				{
					DrawLine(p, x, y - Hang, (x + sz/Scale), y - Hang);  //Font size is not affected by scale.
				}
			}
			else
			{
				FillRectangle(aBrush, x, y -Hang - LineWidth/ 2, sz/Scale, LineWidth);
			}
		}
		#endregion

		#region Interactive Annotations
		/// <summary>
		/// Creates an Hyperlink on the selected region with the specified Url.
		/// </summary>
		/// <param name="x">x coord.</param>
		/// <param name="y">y coord.</param>
		/// <param name="width">Width of the region.</param>
		/// <param name="height">Height of the region.</param>
		/// <param name="url">Url where to navigate to.</param>
		public void Hyperlink(real x, real y, real width, real height, string url)
		{
			Hyperlink(x, y, width, height, new Uri(url)); //we convert to an uri and back to a string to avoid malformed urls (and security holes)
		}

		/// <summary>
		/// Creates an Hyperlink on the selected region with the specified Url.
		/// </summary>
		/// <param name="x">x coord.</param>
		/// <param name="y">y coord.</param>
		/// <param name="width">Width of the region.</param>
		/// <param name="height">Height of the region.</param>
		/// <param name="url">Url where to navigate to.</param>
		public void Hyperlink(real x, real y, real width, real height, Uri url)
		{
			Body.Hyperlink(_tx(x), _thy(y, height), _t(width), _t(height), url.AbsoluteUri);
		}

		/// <summary>
		/// Creates a comment on the pdf file.
		/// </summary>
		/// <param name="x">x coord.</param>
		/// <param name="y">y coord.</param>
		/// <param name="width">Width of the region.</param>
		/// <param name="height">Height of the region.</param>
		/// <param name="comment">Text to put into the comment.</param>
		/// <param name="commentProperties">Properties for the comment.</param>
		public void Comment(real x, real y, real width, real height, string comment, TPdfCommentProperties commentProperties)
		{
			Body.Comment(_tx(x), _thy(y, height), _t(width), _t(height), comment, new TPdfCommentProperties(commentProperties));
		}

		#endregion

		#region Bookmarks
		/// <summary>
		/// Retuns all the bookmarks on the file. Note that this will returna COPY of the bookmarks,
		/// so changing them will not change them in the file. You ned to use <see cref="SetBookmarks"/> to replace the new list.
		/// </summary>
		/// <returns>Existing bookmarks on the file.</returns>
		public TBookmarkList GetBookmarks()
		{
			return (TBookmarkList)Body.Bookmarks.Clone();
		}

		/// <summary>
		/// Replaces the bookmarks on the file with other list. The new list will be copied,
		/// so you can change the old list after setting it and it will not affect the file.
		/// </summary>
		/// <param name="bookmarks">List to replace. If null, bookmarks will be cleared.</param>
		public void SetBookmarks(TBookmarkList bookmarks)
		{
			if (bookmarks == null)
			{
				Body.Bookmarks.Clear();
			}
			else
			{
				Body.Bookmarks = (TBookmarkList) bookmarks.Clone();
			}
		}

		/// <summary>
		/// Adds a new bookmark to the document. 
		/// </summary>
		/// <param name="bookmark"></param>
		public void AddBookmark(TBookmark bookmark)
		{
			Body.Bookmarks.Add(bookmark);
		}

		#endregion

		#region View
		/// <summary>
		/// Sets the default page layout when opening the document.
		/// </summary>
		public TPageLayout PageLayout
		{
			get 	
			{
				return Body.PageLayout;
			}
			set 	
			{
				Body.PageLayout = value;
			}
		}
		#endregion

		#region Measurement
		/// <summary>
		/// For measurements where an actual body might not exist.
		/// </summary>
		/// <returns></returns>
		private TBodySection GetBody()
		{
			if (Body!=null) return Body;
			if (TempBody!= null) return TempBody;
			TempBody = TBodySection.CreateTempBody(FallbackFonts, FontEvents);
			return TempBody;
		}

		/// <summary>
		/// Returns the size of a string in points * Scale. (1/72 of an inch * Scale)
		/// </summary>
		/// <param name="text">String to measure.</param>
		/// <param name="aFont">Font to measure</param>
		/// <returns>Size of the string in points.</returns>
		public SizeF MeasureString( string text, Font aFont)
		{
			if (text == null || text.Length == 0) return SizeF.Empty;
			TPdfFont PdfFont = GetBody().GetFont(FFontMapping, aFont, text, FFontEmbed, FFontSubset, FKerning);
			real sw = ti(PdfFont.MeasureString(text)* aFont.SizeInPoints * 0.001F);
			real sh = ti((PdfFont.LineGap()+Math.Abs(PdfFont.Descent())+Math.Abs(PdfFont.Ascent()))* aFont.SizeInPoints * 0.001F);

			return new SizeF(sw, sh);
		}

		/// <summary>
		/// Returns the font height on points * Scale. (1/72 of an inch * Scale)
		/// </summary>
		/// <param name="aFont">Font to measure.</param>
		/// <returns>Font height.</returns>
		public real FontHeight(Font aFont)
		{
			TPdfFont PdfFont = GetBody().GetFont(FFontMapping, aFont, String.Empty, FFontEmbed, FFontSubset, FKerning);
			return ti((PdfFont.LineGap()+Math.Abs(PdfFont.Descent())+Math.Abs(PdfFont.Ascent()))* aFont.SizeInPoints * 0.001F);
		}

		/// <summary>
		/// Returns the font white space on points * Scale. (1/72 of an inch * Scale)
		/// </summary>
		/// <param name="aFont">Font to measure.</param>
		/// <returns>The blank space between lines.</returns>
		public real FontLinespacing(Font aFont)
		{
			TPdfFont PdfFont = GetBody().GetFont(FFontMapping, aFont, String.Empty, FFontEmbed, FFontSubset, FKerning);
			return ti(PdfFont.LineGap()* aFont.SizeInPoints * 0.001F);
		}

		/// <summary>
		/// Returns the font height on points * Scale. (1/72 of an inch * Scale)
		/// </summary>
		/// <param name="aFont">Font to measure.</param>
		/// <returns>Font height.</returns>
		public real FontDescent(Font aFont)
		{
			TPdfFont PdfFont = GetBody().GetFont(FFontMapping, aFont, String.Empty, FFontEmbed, FFontSubset, FKerning);
			return ti(Math.Abs(PdfFont.Descent())* aFont.SizeInPoints * 0.001F);
		}

		private real FontDescent(TPdfFont aPdfFont, Font aFont)
		{
			return ti(Math.Abs(aPdfFont.Descent())* aFont.SizeInPoints * 0.001F);
		}

		#endregion

		#region Lines and Polygons

		private void MoveTo(real x, real y)
		{
			TPdfBaseRecord.WriteLine(DataStream, tx(x) + " " +
				ty(y) + " " + TPdfTokens.GetString(TPdfToken.CommandMove));
		}

		private void LineTo(real x, real y, TPdfToken Token)
		{
			TPdfBaseRecord.WriteLine(DataStream, tx(x) + " " +
				ty(y) + " " +
				TPdfTokens.GetString(Token));
		}

		private void SplineTo(TPointF p1, TPointF p2, TPointF p3)
		{
			TPdfBaseRecord.WriteLine(DataStream, tx(p1.X) + " " + ty(p1.Y) + " " +
				tx(p2.X) + " " + ty(p2.Y) + " " +
				tx(p3.X) + " " + ty(p3.Y) + " " +
				TPdfTokens.GetString(TPdfToken.CommandBezier));
		}

		private static string GetTokenString(TPdfToken Token)
		{
			if (Token == TPdfToken.None) return String.Empty;
			return TPdfTokens.GetString(Token);
		}

		private void Rectangle(real x, real y, real width, real height, TPdfToken Token)
		{
			TPdfBaseRecord.WriteLine(DataStream, tx(x) + " " +
				thy(y, height) + " " +
				t(width) + " " +
				t(height) + " " +
				TPdfTokens.GetString(TPdfToken.CommandRectangle) + " " +GetTokenString(Token));
		}

		/// <summary>
		/// Draws a line on the current page.
		/// </summary>
		/// <param name="aPen">Pen used to draw the line. Width and color are used.</param>
		/// <param name="x1"></param>
		/// <param name="y1"></param>
		/// <param name="x2"></param>
		/// <param name="y2"></param>
		public void DrawLine(Pen aPen, real x1, real y1, real x2, real y2)
		{
			SelectPen(aPen);
			MoveTo(x1, y1);
			LineTo(x2, y2, TPdfToken.CommandLineToAndStroke);
		}

		/// <summary>
		/// Draws an array of line connecting pairs of points.
		/// </summary>
		/// <param name="aPen">The pen to draw the lines.</param>
		/// <param name="points">An array of points. Its length must be more than 2.</param>
		public void DrawLines(Pen aPen, TPointF[] points)
		{
			SelectPen(aPen);
			if (points.Length < 1) return;

			MoveTo(points[0].X, points[0].Y);
			for (int i = 1; i < points.Length-1; i++)
			{
				LineTo(points[i].X, points[i].Y, TPdfToken.CommandLineTo);
			}

			LineTo(points[points.Length - 1].X, points[points.Length - 1].Y, TPdfToken.CommandLineToAndStroke);
		}

		/// <summary>
		/// Fills the interior of a rectangle specified by a pair of coordinates, a width, and a height.
		/// No line is drawn around the rectangle.
		/// </summary>
		/// <param name="aBrush">Brush to fill.</param>
		/// <param name="x1"></param>
		/// <param name="y1"></param>
		/// <param name="width"></param>
		/// <param name="height"></param>
		public void FillRectangle(Brush aBrush, real x1, real y1, real width, real height)
		{
			if (width == 0 || height == 0) return;

			for (int z = 0; z < BrushIterations(aBrush); z++)
			{
				bool SavesState; SelectBrush(aBrush, z, out SavesState);
				Rectangle(x1, y1, width, height, TPdfToken.CommandFillPath);
				if (SavesState) RestoreState();
			}
		}

		/// <summary>
		/// Draws a rectangle specified by a coordinate pair, a width, and a height. 
		/// </summary>
		/// <param name="aPen">Pen for the line.</param>
		/// <param name="x1"></param>
		/// <param name="y1"></param>
		/// <param name="width"></param>
		/// <param name="height"></param>
		public void DrawRectangle(Pen aPen, real x1, real y1, real width, real height)
		{
			if (width == 0 || height == 0) return;
			SelectPen(aPen);
			Rectangle(x1, y1, width, height, TPdfToken.CommandStroke);
		}

		/// <summary>
		/// Draws and fills a rectangle specified by a coordinate pair, a width, and a height. 
		/// </summary>
		/// <param name="aPen">Pen for the line. Might be null.</param>
		/// <param name="aBrush">Brush for the fill. Might be null.</param>
		/// <param name="x1"></param>
		/// <param name="y1"></param>
		/// <param name="width"></param>
		/// <param name="height"></param>
		public void DrawAndFillRectangle(Pen aPen, Brush aBrush, real x1, real y1, real width, real height)
		{
			if (width == 0 || height == 0) return;
			if (aPen == null && aBrush == null) return;
			SelectPen(aPen);
			for (int z = 0; z < BrushIterations(aBrush); z++)
			{
				bool SavesState; SelectBrush(aBrush, z, out SavesState);
				try
				{
					Rectangle(x1, y1, width, height, TPdfToken.None);

					if (aPen != null && aBrush != null && !SavesState)
					{
						TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandFillAndStroke));
					}
					else
					{
						if (aBrush != null)
						{
							TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandFillPath));
							if (SavesState) 
							{ 
								RestoreState(); SavesState = false; 
								if (aPen != null) Rectangle(x1, y1, width, height, TPdfToken.None); //allow for the next command to continue.
							}
						}
						if (aPen != null)
						{
							TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandStroke));
						}
					}
				}
				finally
				{
					if (SavesState) RestoreState();
				}
			}
		}

		/// <summary>
		/// Draws and/or fills a bezier path. If Brush is not null, the shape will be closed for filling.
		/// </summary>
		/// <param name="aPen">Pen for the outline. If null, no outline will be drawn.</param>
		/// <param name="aBrush">Brush for filling. If null, the shape will not be filled.</param>
		/// <param name="aPoints">Array of points for the curve. See GDI+ DrawBeziers function for more information.</param>
		public void DrawAndFillBeziers(Pen aPen, Brush aBrush, TPointF[] aPoints)
		{
			if (aPen == null && aBrush == null) return;
			if (aPoints == null || aPoints.Length <= 0) return;
			SelectPen(aPen);

			for (int z = 0; z < BrushIterations(aBrush); z++)
			{
				bool SavesState; SelectBrush(aBrush, z, out SavesState);
				try
				{
					DrawBezierPath(aPoints);

					if (aPen != null && aBrush != null && !SavesState)
					{
						TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandFillAndStroke));
					}
					else
					{
						if (aBrush != null)
						{
							TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandFillPath));
							if (SavesState) { RestoreState(); SavesState = false; DrawBezierPath(aPoints); }
						}
						if (aPen != null)
						{
							TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandStroke));
						}
					}
				}
				finally
				{
					if (SavesState) RestoreState();
				}
			}
		}

		private void DrawBezierPath(TPointF[] aPoints)
		{
			MoveTo(aPoints[0].X, aPoints[0].Y);
			int i = 1;
			while (i + 2 < aPoints.Length)
			{
				SplineTo(aPoints[i], aPoints[i + 1], aPoints[i + 2]);
				i += 3;
			}
		}

		/// <summary>
		/// Draws and/or fills a polygon. The shape will be closed.
		/// </summary>
		/// <param name="aPen">Pen for the outline. If null, no outline will be drawn.</param>
		/// <param name="aBrush">Brush for filling. If null, the shape will not be filled.</param>
		/// <param name="aPoints">Array of points for the polygon.</param>
		public void DrawAndFillPolygon(Pen aPen, Brush aBrush, TPointF[] aPoints)
		{
			if (aPoints == null || aPoints.Length <= 0) return;
			SelectPen(aPen);
			for (int z = 0; z < BrushIterations(aBrush); z++)
			{
				bool SavesState; SelectBrush(aBrush, z, out SavesState);
				try
				{
					DrawPolygonPath(aPoints);

					if (aPen != null && aBrush != null && !SavesState)
					{
						TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandFillAndStroke));
					}
					else
					{
						if (aBrush != null)
						{
							TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandFillPath));
							if (SavesState) { RestoreState(); SavesState = false; DrawPolygonPath(aPoints);}
						}
						if (aPen != null)
						{
							TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandClosePath)+ " "+ TPdfTokens.GetString(TPdfToken.CommandStroke));
						}
					}
				}
				finally
				{
					if (SavesState) RestoreState();
				}
			}
		}

		private void DrawPolygonPath(TPointF[] aPoints)
		{
			MoveTo(aPoints[0].X, aPoints[0].Y);
			for (int i = 1; i < aPoints.Length; i++)
			{
				LineTo(aPoints[i].X, aPoints[i].Y, TPdfToken.CommandLineTo);
			}

		}

		#endregion

		#region Clipping
		/// <summary>
		/// Intersects the current clipping region with the new one. 
		/// There is no command to reset or expand a clipping region, you need to use
		/// <see cref="SaveState"/> and <see cref="RestoreState"/>
		/// </summary>
		/// <param name="Rect"></param>
		public void IntersectClipRegion(RectangleF Rect)
		{
			Rectangle(Rect.Left, Rect.Top, Rect.Width, Rect.Height, TPdfToken.CommandClipPath);
		}

		/// <summary>
		/// Intersect the clip region with a rectangle specified by a pair of coordinates, a width, and a height.
		/// </summary>
		/// <param name="x1"></param>
		/// <param name="y1"></param>
		/// <param name="width"></param>
		/// <param name="height"></param>
		/// <param name="exclude">When true, all region OUTSIDE the rectangle will be intersected with the current clipping region.</param>
		public void ClipRectangle(real x1, real y1, real width, real height, bool exclude)
		{
			if (exclude)
			{
				Rectangle(0, 0, PageSize.Width, PageSize.Height, TPdfToken.None);
			}
			Rectangle(x1, y1, width, height, TPdfToken.CommandClipPathEvenOddRule); 
		}

		/// <summary>
		/// Intersects the clipping area with a polygon. 
		/// </summary>
		/// <param name="aPoints">Array of points for the polygon.</param>
		/// <param name="exclude">When true, all region OUTSIDE the polygon will be intersected with the current clipping region.</param>
		public void ClipPolygon(TPointF[] aPoints, bool exclude)
		{
			if (aPoints == null || aPoints.Length <= 0) return;
			if (exclude)
			{
				Rectangle(0, 0, PageSize.Width, PageSize.Height, TPdfToken.None);
			}

			MoveTo(aPoints[0].X, aPoints[0].Y);
			for (int i = 1; i < aPoints.Length; i++)
			{
				LineTo(aPoints[i].X, aPoints[i].Y, TPdfToken.CommandLineTo);
			}

			TPdfBaseRecord.Write(DataStream, TPdfTokens.GetString(TPdfToken.CommandClosePath)+" ");
			TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandClipPathEvenOddRule));
		}

		/// <summary>
		/// Intersects the clipping area with a bezier region. 
		/// </summary>
		/// <param name="aPoints">Array of points for the curve. See GDI+ DrawBeziers function for more information.</param>
		/// <param name="exclude">When true, all region OUTSIDE the region will be intersected with the current clipping region.</param>
		public void ClipBeziers(TPointF[] aPoints, bool exclude)
		{
			if (aPoints == null || aPoints.Length <= 0) return;
			if (exclude)
			{
				Rectangle(0, 0, PageSize.Width, PageSize.Height, TPdfToken.None);
			}

			MoveTo(aPoints[0].X, aPoints[0].Y);
			int i = 1;
			while (i+2 < aPoints.Length)
			{
				SplineTo(aPoints[i], aPoints[i+1], aPoints[i+2]);
				i+=3;
			}

			TPdfBaseRecord.Write(DataStream, TPdfTokens.GetString(TPdfToken.CommandClosePath)+" ");
			TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CommandClipPathEvenOddRule));
		}

		#endregion

		#region Images
		/// <summary>
		/// Draws an image on the canvas. Image can be any type, but all except jpeg will be converted to PNG.
		/// </summary>
		/// <param name="image">Image to insert. If it is not JPEG or PNG, it will be converted to PNG.</param>
		/// <param name="rect">Rectangle where the image will be.</param>
		/// <param name="imageData">Stream with the raw image data. Not required, might be null.</param>
		public void DrawImage(Image image, RectangleF rect, Stream imageData)
		{
			DrawImage(image, rect, imageData, ~0L, false);
		}

		/// <summary>
		/// Draws an image on the canvas. Image can be any type, but all except jpeg will be converted to PNG.
		/// </summary>
		/// <param name="image">Image to insert. If it is not JPEG or PNG, it will be converted to PNG.</param>
		/// <param name="rect">Rectangle where the image will be.</param>
		/// <param name="imageData">Stream with the raw image data. Not required, might be null.</param>
		/// <param name="transparentColor">Color to make transparent. To specify no transparent color use <see cref="FlexCel.Core.FlxConsts.NoTransparentColor"/></param>
		/// <param name="defaultToJpg">When true and the image is not on a supported format (or imageData==null) the image will
		/// be converted to JPG. If false, the image will be converted to PNG.</param>
		public void DrawImage(Image image, RectangleF rect, Stream imageData, long transparentColor, bool defaultToJpg)
		{
			if (image.Width <= 0 || image.Height <= 0) return;
			if (rect.Width <= 0 || rect.Height <= 0) return;
			ResetBrushTransparency();

			SaveState();
			TPdfBaseRecord.WriteLine(DataStream, String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} {4} {5} cm", 
				t(rect.Width), 
				0, 0,
				t(rect.Height),
				tx(rect.Left), 
				thy(rect.Top, rect.Height)));
			Body.SelectImage(DataStream, image, imageData, transparentColor, defaultToJpg);
			RestoreState();
		}

		#endregion

		#region Rotate And Scale
		/// <summary>
		/// Rotates the canvas around point (x,y).
		/// </summary>
		/// <param name="x">Point where the canvas is rotated.</param>
		/// <param name="y">Point where the canvas is rotated.</param>
		/// <param name="Alpha">Angle in degrees.</param>
		public void Rotate(real x, real y, real Alpha)
		{
			double Ar = Alpha * Math.PI / 180;
			double CosAlpha = Math.Cos(Ar);
			double SinAlpha = Math.Sin(Ar);
			String CosAlphaS= PdfConv.DoubleToString(CosAlpha);
			String SinAlphaS= PdfConv.DoubleToString(SinAlpha);
			String NegSinAlphaS= PdfConv.DoubleToString(-SinAlpha);
			TPdfBaseRecord.WriteLine(DataStream, String.Format("{0} {1} {2} {3} {4} {5} cm", 1, 0, 0, 1, t(x), ty(y)));
			TPdfBaseRecord.WriteLine(DataStream, String.Format("{0} {1} {2} {3} {4} {5} cm", CosAlphaS, SinAlphaS, NegSinAlphaS, CosAlphaS, 0, 0));
			TPdfBaseRecord.WriteLine(DataStream, String.Format("{0} {1} {2} {3} {4} {5} cm", 1, 0, 0, 1, t(-x), tnegy(y)));

			real negy = _tnegy(y); //It will change when translated.
			TranslateMatrix(_t(x), _ty(y));
			RotateMatrix(CosAlpha, SinAlpha);
			TranslateMatrix( _t(-x), negy);

		}

		/// <summary>
		/// Scales the canvas. It premultiplies the matrix, to keep the correct order.
		/// </summary>
		/// <param name="xScale">X Scale</param>
		/// <param name="yScale">Y Scale</param>
		public void ScaleBy(real xScale, real yScale)
		{
			ScaleMatrix(xScale, yScale);
			TPdfBaseRecord.WriteLine(DataStream, String.Format("{0} {1} {2} {3} {4} {5} cm", PdfConv.DoubleToString(xScale), 0, 0, PdfConv.DoubleToString(yScale), 0, 0));
		}

		/// <summary>
		/// Returns the drawing matrix in use. The elements in this matrix are similar to the ones returned by
		/// <see cref="System.Drawing.Drawing2D.Matrix.Elements"/> and have the same meaning.
		/// <b>Important remark. This matrix is the real one, and does not consider things like <see cref="YAxisGrowsDown"/> or <see cref="Scale"/>.</b>
		/// You will probably want to use <see cref="Transform(TPointF)"/> to find out the coordinates of a point.
		/// </summary>
		/// <returns>The internal transformation matrix.</returns>
		public double[] GetMatrix()
		{
			return (double[]) DrawingMatrix.Clone();
		}

		/// <summary>
		/// Transforms the point according to the current transformation Matrix. See <see cref="GetMatrix"/> to get the actual matrix.
		/// </summary>
		/// <param name="p">Point you want to map to the user coordinates.</param>
		/// <returns></returns>
		public TPointF Transform(TPointF p)
		{
            return new TPointF(
				ti((real)(DrawingMatrix[0] * _tx(p.X) + DrawingMatrix[2] * _ty(p.Y) + DrawingMatrix[4])),
				tinvy((real)(DrawingMatrix[1] * _tx(p.X) + DrawingMatrix[3] * _ty(p.Y) + DrawingMatrix[5])));
		}

		#endregion

		#region Graphics state
		/// <summary>
		/// Saves the current graphic state. Be sure to always call <see cref="RestoreState"/>
		/// each time you call this method.
		/// </summary>
		public void SaveState()
		{
			GraphicsState.Push(new TGState(LastBrushColor, LastBrushStyle, LastHatchStyle, LastPenColor, LastPenWidth, LastPenStyle, DrawingMatrix));
			TPdfBaseRecord.WriteLine(DataStream, "q");
		}

		/// <summary>
		/// Restores the graphic state saved by a <see cref="SaveState"/> call.
		/// </summary>
		public void RestoreState()
		{
			TGState St = GraphicsState.Pop();
			TPdfBaseRecord.WriteLine(DataStream, "Q");
			LastBrushColor = St.LastBrushColor;
			LastHatchStyle = St.LastHatchStyle;
			LastBrushStyle = St.LastBrushStyle;
			LastPenColor = St.LastPenColor;
			LastPenWidth = St.LastPenWidth;
			LastPenStyle = St.LastPenStyle;
			DrawingMatrix = (double[])St.DrawingMatrix.Clone();
		}

		#endregion

		#region Utilities
		private real CurrentYScale
		{
			get
			{
				/* x' = a11x + a21y + a31.  x=0 (we are converting PageHeight) and a31 = 0, so x' = a21y
				 * y' = a22y
				 * r2 = x'2 + y'2 = (a21^2 + a22^2) * y2
				 */
                 
				double r = Math.Sqrt(DrawingMatrix[2] * DrawingMatrix[2] +DrawingMatrix[3] * DrawingMatrix[3]);
				if (r == 0) return 0;
				return (real) (1/r);
			}
		}

		private void TranslateMatrix(double dx, double dy)
		{
			DrawingMatrix[4] += dx * DrawingMatrix[0] + dy * DrawingMatrix[2];
			DrawingMatrix[5] += dx * DrawingMatrix[1] + dy * DrawingMatrix[3];
		}

		private void ScaleMatrix(double xScale, double yScale)
		{
			DrawingMatrix[0] *= xScale;
			DrawingMatrix[1] *= xScale;
			DrawingMatrix[2] *= yScale;
			DrawingMatrix[3] *= yScale;
		}

		private void RotateMatrix(double CosAlpha, double SinAlpha)
		{
			double Result;
			for (int i = 0; i < 2; i++)
			{
				Result = (DrawingMatrix[0 + i] * CosAlpha + DrawingMatrix[2 + i] * SinAlpha);
				DrawingMatrix[2 + i] = (-DrawingMatrix[0 + i] * SinAlpha + DrawingMatrix[2 + i] * CosAlpha);

				DrawingMatrix[0 + i] = Result;
			}
		}

		private RectangleF GetPdfRectangle(RectangleF Source)
		{
			return new RectangleF(_tx(Source.X), _thy(Source.Y, Source.Height), _t(Source.Width), _t(Source.Height));
		}

		private static int BrushIterations(Brush aBrush)
		{
			if (aBrush is HatchBrush) return 2;
			return 1;
		}

		private bool BrushIsTransparent(Brush aBrush)
		{
			SolidBrush sb = aBrush as SolidBrush;
			if (sb != null) return sb.Color.A != LastBrushColor.A;
			return true; //This method is not complete, and it does not need to be. It just needs to return true if in doubt.
		}

        private static RectangleF CalcRotatedCoords(float[] aPatternMatrix, RectangleF aBrushCoords)
        {
            //The definition in PDF is not the containing rect (which we could find by applying the transform to the original rect), 
            //but a line that is perpendicular to the gradient. Also, this is rotated -90 degrees from the original Excel def.
            //So, the algorithm is:
            //1) Take the line (xcenter, top)(xcenter,bottom)
            //2) Transform this line with the transform matrix.
            //3) Rotate the line -90 degrees to find the perpendicular to that.

            double w2 = aBrushCoords.Width / 2.0;
            double h2 = aBrushCoords.Height / 2.0;

            double xCenter = aBrushCoords.Left + w2;
            //double yCenter = aBrushCoords.Top + h2;

            double x0 = xCenter * aPatternMatrix[0] + aBrushCoords.Y * aPatternMatrix[2] + aPatternMatrix[4];
            double y0 = xCenter * aPatternMatrix[1] + aBrushCoords.Y * aPatternMatrix[3] + aPatternMatrix[5];

            double x1 = xCenter * aPatternMatrix[0] + aBrushCoords.Bottom * aPatternMatrix[2] + aPatternMatrix[4];
            double y1 = xCenter * aPatternMatrix[1] + aBrushCoords.Bottom * aPatternMatrix[3] + aPatternMatrix[5];

            //Excel 2007 is buggy in this, we will follow Excel 2003. Thing is, when you define for example an angle of 45 degrees and
            //width of the box is twice the height, the "real" angle of the gradient is not 45. Now, in Excel 2007 sometimes it is, sometimes it isn't.

            //We got the line that goes along with the gradient. But to define it in pdf, we need to find the perpendicular.
            //so we will rotate it 90 degrees against the center. 

            double dx = Math.Abs(x1 - x0) / 2;
            double dy = Math.Abs(y1 - y0) / 2;
            double dx2 = dx*dx;
            double dy2 = dy*dy;
            double dxy = dx*dy;
            double x2Plusy2 = dx2 + dy2;

            if (x2Plusy2 == 0) x2Plusy2 = 1; //doesn't really matter, means width and height are 0.

            double ndx = (w2 * dy2 + h2 * dxy) / x2Plusy2;
            double ndy = (w2 * dxy + h2 * dx2) / x2Plusy2;

            double nxCenter = (x0 + x1)/2;
            double nyCenter = (y0 + y1)/2;

            if (y1 < y0) ndx = -ndx;
            if (x1 < x0) ndy = -ndy;

            return RectangleF.FromLTRB((real)(nxCenter + ndx), (real)(nyCenter + ndy), (real)(nxCenter - ndx), (real)(nyCenter - ndy));
        }

		private void SelectBrush(Brush aBrush, int Pass, out bool SavedState)
		{
			//Note that instead of the "SaveState" approach of saving state for gradients, we could have specified a "None" SMask to revert its effects. It might be a little cleaner (we will not have to have a SaveState var) but also more confusing. (we would need to check for smask and set it to none on the solid brushes)
			SavedState = false;
			if (aBrush == null) return;
			HatchBrush HBrush = (aBrush as HatchBrush);

			if (Pass == 0 && HBrush != null)
			{
				aBrush = new SolidBrush(HBrush.BackgroundColor);
				HBrush = null;
			}

			if (HBrush != null)
			{
				//Here we will compare with Color.Equals and not color.ToArgb. Because Color.Equals returns true only if both Color entities are the same, including the name. Since LastBrushColor is Color.Black at the beginning, it must be different from ColorUtil.FromArgb(1,0,0,0);
				if (LastBrushStyle != TBrushStyle.Hatch || (!Color.Equals(HBrush.ForegroundColor, LastBrushColor)) || LastHatchStyle != HBrush.HatchStyle)
				{
					Body.SelectBrush(DataStream, HBrush);
	
					if (LastBrushColor.A != HBrush.ForegroundColor.A) //GState is shared between Hatch brushes and solid brushes. So we do not have to check for different variables here.
					{
						Body.SelectTransparency(DataStream, HBrush.ForegroundColor.A, TPdfToken.CommandSetAlphaBrush);
					}

					LastBrushStyle = TBrushStyle.Hatch;
					LastHatchStyle = HBrush.HatchStyle;
					LastBrushColor = HBrush.ForegroundColor;

				}
				return;
			}
            
			SolidBrush MyBrush = (aBrush as SolidBrush);
			if (MyBrush != null)
			{
				//Here we will compare with Color.Equals and not color.ToArgb. Because Color.Equals returns true only if both Color entities are the same, including the name. Since LastBrushColor is Color.Black at the beginning, it must be different from ColorUtil.FromArgb(1,0,0,0);
				if(LastBrushStyle != TBrushStyle.Solid || (!Color.Equals(MyBrush.Color, LastBrushColor)))
				{
					TPdfToken AlphaToken = TPdfToken.None;
					if (LastBrushColor.A != MyBrush.Color.A) //GState is shared between Hatch brushes and solid brushes. So we do not have to check for different variables here.
					{
						AlphaToken = TPdfToken.CommandSetAlphaBrush;
					}
					LastBrushStyle = TBrushStyle.Solid;
					LastHatchStyle = (HatchStyle) (-1);
					LastBrushColor = MyBrush.Color;
                
					SelectColor(LastBrushColor, TPdfToken.CommandSetBrushColor, AlphaToken);
				}
				return;
			}

            
			LinearGradientBrush LGBrush = (aBrush as LinearGradientBrush);
			if (LGBrush != null)
			{
				ResetBrushTransparency();
				RectangleF Coords = GetPdfRectangle(LGBrush.Rectangle);
                RectangleF RotatedCoords = GetPdfRectangle(CalcRotatedCoords(LGBrush.Transform.Elements, LGBrush.Rectangle));
				SaveState(); SavedState = true;
				Body.SelectBrush(DataStream, LGBrush, Coords, RotatedCoords, PdfConv.ToString(DrawingMatrix, true));

				LastBrushStyle = TBrushStyle.GradientLinear; //it doesn't matter. It will be restored by RestoreState
				LastHatchStyle = (HatchStyle) (-1);
				return;
			}

			PathGradientBrush PGBrush = (aBrush as PathGradientBrush);
			if (PGBrush != null)
			{
				ResetBrushTransparency();
				RectangleF Coords = GetPdfRectangle(PGBrush.Rectangle);
                RectangleF RotatedCoords = GetPdfRectangle(CalcRotatedCoords(PGBrush.Transform.Elements, PGBrush.Rectangle));
                PointF CenterPoint = new PointF(_tx(PGBrush.CenterPoint.X), _ty(PGBrush.CenterPoint.Y));
				SaveState(); SavedState = true;
				Body.SelectBrush(DataStream, PGBrush, Coords, RotatedCoords, CenterPoint, PdfConv.ToString(DrawingMatrix, true));

				LastBrushStyle = TBrushStyle.GradientRadial;  //it doesn't matter. It will be restored by RestoreState
				LastHatchStyle = (HatchStyle) (-1);
				return;
			}

			TextureBrush TBrush = aBrush as TextureBrush;
			if (TBrush != null)
			{
				ResetBrushTransparency();
				real[] PatternMatrix = TBrush.Transform.Elements;
				PatternMatrix[4] = _tx(PatternMatrix[4]);
				PatternMatrix[5] = _ty(PatternMatrix[5]);

				Body.SelectBrush(DataStream, TBrush, PatternMatrix);
	
				LastBrushStyle = TBrushStyle.Texture;
				LastHatchStyle = (HatchStyle) (-1);
				return;
			}
		}

        private void SelectPen(Pen aPen)
        {
            if (aPen == null) return;
			//Here we will compare with Color.Equals and not color.ToArgb. Because Color.Equals returns true only if both Color entities are the same, including the name. Since LastBrushColor is Color.Black at the beginning, it must be different from ColorUtil.FromArgb(1,0,0,0);
			if (!Color.Equals(aPen.Color,  LastPenColor))
            {
                TPdfToken AlphaToken = TPdfToken.None;
                if (LastPenColor.A != aPen.Color.A)
                {
                    AlphaToken = TPdfToken.CommandSetAlphaPen;
                }
                LastPenColor = aPen.Color;
                SelectColor(LastPenColor, TPdfToken.CommandSetPenColor, AlphaToken);
            }
            if (LastPenWidth < 0 || aPen.Width != LastPenWidth)
            {
                LastPenWidth = aPen.Width;
                TPdfBaseRecord.WriteLine(DataStream, t(LastPenWidth) + " " + TPdfTokens.GetString(TPdfToken.CommandLineWidth));
            }

            if (LastPenStyle != aPen.DashStyle)
            {
                LastPenStyle = aPen.DashStyle;
                SelectLineStyle(LastPenStyle);
            }

        }

        private void SelectColor(Color aColor, TPdfToken Token, TPdfToken AlphaToken)
        {
            TPdfBaseRecord.WriteLine(DataStream,
                PdfConv.DoubleToString(aColor.R / 255.0) + " " +
                PdfConv.DoubleToString(aColor.G / 255.0) + " " +
                PdfConv.DoubleToString(aColor.B / 255.0) + " " +
                TPdfTokens.GetString(Token));

            if (AlphaToken != TPdfToken.None)
            {
                Body.SelectTransparency(DataStream, aColor.A, AlphaToken);
            }

        }

        private static string GetDashStyle(DashStyle ds)
        {
            switch (ds)
            {
                case DashStyles.Dot: return "1";
                case DashStyles.Dash: return "2 1";
                case DashStyles.DashDot: return "4 2 2 2";
                case DashStyles.DashDotDot: return "4 2 2 2 2 2";
                default: return String.Empty;
            }
        }

        private void SelectLineStyle(DashStyle ds)
        {
            TPdfBaseRecord.WriteLine(DataStream,
                TPdfTokens.GetString(TPdfToken.OpenArray) +
                GetDashStyle(ds) +
                TPdfTokens.GetString(TPdfToken.CloseArray)+

                " 0 " +
                TPdfTokens.GetString(TPdfToken.CommandSetLineStyle));
        }
		
		private void ResetBrushTransparency()
		{
			//We cannot just restore/save the state here, since a SelectBrush might be a command in a queue. If we restore the original state, for example a SelectPen inmediately before a SelectBrush would be erased.
			if (LastBrushColor.A == 255) return;
			Body.SelectTransparency(DataStream, 255, TPdfToken.CommandSetAlphaBrush);
			LastBrushColor = ColorUtil.FromArgb(255, LastBrushColor.R, LastBrushColor.G, LastBrushColor.B);
		}
        #endregion

		#region Sign and Encrypt
		/// <summary>
		/// Signs the pdf documents with the specified <see cref="TPdfSignature"/> or <see cref="TPdfVisibleSignature"/>.
		/// <b>Note:</b> This method must be called <b>before</b> calling <see cref="BeginDoc"/>
		/// </summary>
		/// <param name="aSignature">Signature used for signing. Set it to null to stop signing the next documents.</param>
		public void Sign(TPdfSignature aSignature)
		{
			if (Body != null) PdfMessages.ThrowException(PdfErr.ErrTryingToSignStartedDocument);
			FSignature = aSignature;
		}
		#endregion

    }

    class TGState
    {
        internal Color LastBrushColor;
        internal TBrushStyle LastBrushStyle;
        internal HatchStyle LastHatchStyle;

        internal Color LastPenColor;
        internal real LastPenWidth;
        internal DashStyle LastPenStyle;

        internal double[] DrawingMatrix;

        internal TGState(Color aLastBrushColor, TBrushStyle aLastBrushStyle, HatchStyle aLastHatchStyle, Color aLastPenColor, real aLastPenWidth, DashStyle aLastPenStyle, double[] aDrawingMatrix)
        {
            LastBrushColor = aLastBrushColor;
            LastHatchStyle = aLastHatchStyle;

            LastPenColor = aLastPenColor;
            LastPenWidth = aLastPenWidth;
            LastBrushStyle = aLastBrushStyle;
            LastPenStyle = aLastPenStyle;
            DrawingMatrix = (double[])aDrawingMatrix.Clone();
        }

    }


    internal enum TBrushStyle
    {
        None,
        Solid,
        Hatch,
        GradientLinear,
        GradientRadial,
		Texture
    }
}
