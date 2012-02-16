using System;
using FlexCel.Core;

#if (WPF)
#else
using System.Drawing;
#endif

namespace FlexCel.Pdf
{
	#region Event Handlers
	/// <summary>
	/// Arguments passed on <see cref="FlexCel.Render.FlexCelPdfExport.GetFontData"/>.
	/// Use this event to provide font information for embedding.
	/// </summary>
	public class GetFontDataEventArgs: EventArgs
	{
		private bool FApplied;
		private byte[] FFontData;
		private readonly Font FFont;

		/// <summary>
		/// Creates a new Argument.
		/// </summary>
		public GetFontDataEventArgs(Font aFont)
		{
			FApplied = true;
			FFont = aFont;
		}

		/// <summary>
		/// Return the full font file as a byte array here.
		/// </summary>
		public byte[] FontData
		{
			get {return FFontData;}
			set {FFontData = value;}
		}

		/// <summary>
		/// Set Applied = false if the font is not being processed by the event.
		/// </summary>
		public bool Applied {get{return FApplied;} set {FApplied = value;}}

		/// <summary>
		/// The font for which you need to return the data.
		/// </summary>
		public Font InputFont {get {return FFont;}}
	}

	/// <summary>
	/// Delegate for reading the font data.
	/// </summary>
	public delegate void GetFontDataEventHandler(object sender, GetFontDataEventArgs e);

	/////////////////////////////////////////////////////////////
	
	
	/// <summary>
	/// Arguments passed on <see cref="FlexCel.Render.FlexCelPdfExport.GetFontFolder"/>.
	/// Use this event to provide font information for embedding.
	/// </summary>
	public class GetFontFolderEventArgs: EventArgs
	{
		private bool FApplied;
        private bool FAppendFontFolder;
        private string FFontFolder;
		private readonly Font FFont;

		/// <summary>
		/// Creates a new Argument.
		/// </summary>
		public GetFontFolderEventArgs(Font aFont, string aPath)
		{
			FApplied = true;
            FAppendFontFolder = true;
			FFont = aFont;
			FFontFolder = aPath;
		}

		/// <summary>
		/// Return here the font path to the "Fonts" folder where ttf files are located. <br></br>When <see cref="AppendFontFolder"/> is true,
        /// FlexCel will add a "..\Fonts" string to what you specify here. For example, if you specify "c:\Windows\System", FlexCel will look at
        /// "c:\Windows\System\..\Fonts". If <see cref="AppendFontFolder"/> is false, the string here will be used literally. If you specify "c:\Windows\Fonts",
        /// FlexCel will look in "c:\Windows\Fonts".
		/// </summary>
		public string FontFolder
		{
			get {return FFontFolder;}
			set {FFontFolder = value;}
		}

		/// <summary>
		/// Set Applied = false if the font is not being processed by the event, and FlexCel should try to find the font path for the font as if the event wasn't assigned.
		/// </summary>
		public bool Applied {get{return FApplied;} set {FApplied = value;}}

        /// <summary>
        /// When true (the default) FlexCel will append "..\Fonts" to the end of the string you specify in <see cref="FontFolder"/>, for backward compatibility 
        /// reasons. If you set this to false, FlexCel won't append anything and just use whatever you write in <see cref="FontFolder"/> property.
        /// For new applications, it is recommended to set this property to false.
        /// </summary>
        public bool AppendFontFolder { get { return FAppendFontFolder; } set { FAppendFontFolder = value; } }

		/// <summary>
		/// The font for which you need to return the data.
		/// </summary>
		public Font InputFont {get {return FFont;}}
	}

	/// <summary>
	/// Delegate for reading the font data.
	/// </summary>
	public delegate void GetFontFolderEventHandler(object sender, GetFontFolderEventArgs e);

	/////////////////////////////////////////////////////////////

	/// <summary>
	/// Arguments passed on <see cref="FlexCel.Render.FlexCelPdfExport.OnFontEmbed"/>.
	/// Use this event to specify which fonts to embed and which fonts to ignore. Note that unicode fonts will be 
	/// embedded no matter what you say here.
	/// </summary>
	public class FontEmbedEventArgs: EventArgs
	{
		private Font FInputFont;
		private bool FEmbed;

		/// <summary>
		/// Creates a new Argument.
		/// </summary>
		public FontEmbedEventArgs(Font aInputFont, bool aEmbed)
		{
			FInputFont = aInputFont;
			FEmbed = aEmbed;
		}

		/// <summary>
		/// The font for which you need to return the data.
		/// </summary>
		public Font InputFont {get {return FInputFont;}}

		/// <summary>
		/// Return true if you want to embed this font, false if you don't want to. If you don't modify this value, the default will be used.
		/// </summary>
		public bool Embed {get {return FEmbed;} set{FEmbed = value;}}
	}

	/// <summary>
	/// Delegate for reading the font data.
	/// </summary>
	[System.Diagnostics.CodeAnalysis.SuppressMessage ("Microsoft.Design", "CA1003:UseGenericEventHandlerInstances")]
    public delegate void FontEmbedEventHandler(object sender, FontEmbedEventArgs e);


	#endregion

	#region Font events
	internal class TFontEvents
	{
		internal object Sender;
		internal GetFontDataEventHandler OnGetFontData;
		internal GetFontFolderEventHandler OnGetFontFolder;
		internal FontEmbedEventHandler OnFontEmbed;

		internal TFontEvents(object aSender, GetFontDataEventHandler aOnGetFontData, GetFontFolderEventHandler aOnGetFontFolder, FontEmbedEventHandler aOnFontEmbed)
		{
			Sender = aSender;
			OnGetFontData = aOnGetFontData;
			OnGetFontFolder = aOnGetFontFolder;
			OnFontEmbed = aOnFontEmbed;
		}
	}
	#endregion
}
