#region Using directives

using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using System.Globalization;

#if (WPF)
using real = System.Double;
using System.Windows;
#else
using real = System.Single;
using Colors = System.Drawing.Color;
using System.Drawing;
#endif

using FlexCel.Core;
using System.Security;
using System.Runtime.CompilerServices;


#endregion

namespace FlexCel.Pdf
{

#if(FRAMEWORK20)
    internal sealed class TPdfFontList : Dictionary<string, TPdfFont>
    {
		internal List<string> OrderedKeys;

        public TPdfFontList(): base()
        {
        }

	    internal void CreateOrderedKeys()
		{
			OrderedKeys = new List<string>();
		}

    }
#else
	internal class TPdfFontList: Hashtable
	{
		internal ArrayList OrderedKeys;

		internal void CreateOrderedKeys()
		{
			OrderedKeys = new ArrayList();
		}
	}

#endif



	internal class TPdfEmbeddedFont
	{
		private int ObjectId;
		private bool AlreadySaved;

		internal TGlyphMap SubsetCMap;
		internal readonly TPdfTrueType TrueTypeData;

		internal TPdfEmbeddedFont(bool Subset, TPdfTrueType aTrueTypeData)
		{
			ObjectId = -1;
			if (Subset) 
			{
				SubsetCMap = new TGlyphMap();
				SubsetCMap.Add(0,0); //Standard missing glyph.
			}
			TrueTypeData = aTrueTypeData;
			AlreadySaved = false;
		}


		internal int GetStreamId(TPdfStream DataStream, TXRefSection XRef)
		{
			if (ObjectId == -1)
			{
				ObjectId = XRef.GetNewObject(DataStream);
			}

			return ObjectId;
		}

		internal void WriteFont(TPdfStream DataStream, TXRefSection XRef, bool Compress)
		{
			if (AlreadySaved || ObjectId == -1) return;
			AlreadySaved = true;
			int EmbeddedLenId = XRef.GetNewObject(DataStream); //Stream Length

			XRef.SetObjectOffset(ObjectId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, ObjectId);
				
			byte[] FontData = TrueTypeData.SubsetFontData(SubsetCMap);

			TStreamRecord.BeginSave(DataStream, EmbeddedLenId, Compress, FontData.Length);
			long StartStream = DataStream.Position;
			bool Compressing = DataStream.Compress;
			try
			{
				DataStream.Compress = Compress;
				DataStream.Write(FontData);
			}
			finally
			{
				DataStream.Compress = Compressing;
			}
			long EndStream = DataStream.Position;
			TStreamRecord.EndSave(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);

			XRef.SetObjectOffset(EmbeddedLenId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, EmbeddedLenId);
			TPdfBaseRecord.WriteLine(DataStream, (EndStream - StartStream).ToString(CultureInfo.InvariantCulture));
			TIndirectRecord.SaveTrailer(DataStream);
		}

		internal void AddGlyphFromChar(int UnicodeChar)
		{
			if (SubsetCMap == null) return;

			int gl = TrueTypeData.Glyph(UnicodeChar, true);
			if (!SubsetCMap.ContainsKey(gl)) SubsetCMap.Add(gl, SubsetCMap.Count);
		}

		internal int GetNewGlyphFromOldGlyph(int glCode)
		{
			if (SubsetCMap == null) return glCode;

			int gl;
			if (SubsetCMap.TryGetValue(glCode, out gl)) return gl;

			int NewGlyph = SubsetCMap.Count;
			SubsetCMap.Add(glCode, NewGlyph);
			return NewGlyph;
		}

	}

    internal sealed class TPdfEmbeddedFontList : Dictionary<string, TPdfEmbeddedFont>
    {

        private static bool WithoutGetFolderPathPermissions = false;
#if (FRAMEWORK30 && !WPF && !__MonoCS__ && COMMENTED_OUT_BECAUSE_IT_CAN_FAIL_IN_MULTITHREAD_CASES)
        private static bool WithoutWPFSupport = false;
#endif

        public TPdfEmbeddedFont Add(Font aFont, TFontEvents FontEvents, bool Subset, bool UseKerning, out FontStyle AdditionalStyle)
		{
			TPdfEmbeddedFont Result;
			//We can't search first for the font name, because 2 fonts with the same name might have different font files. (one might be a font file for italic or bold)
			//if (TryGetValue(aFont.Name, out Result)) return Result;

			byte[] FontData = LoadFont(aFont, FontEvents);
		
			TPdfTrueType TrueTypeData = new TPdfTrueType(FontData, aFont.Name, UseKerning);

            AdditionalStyle = FontStyle.Regular;
            FontStyle fs = TPsFontList.GetStyle(TrueTypeData.FontFlags);
            if (aFont.Italic && ((fs & FontStyle.Italic) == 0)) AdditionalStyle |= FontStyle.Italic;
            if (aFont.Bold && ((fs & FontStyle.Bold) == 0)) AdditionalStyle |= FontStyle.Bold;
			string FontFileName = TrueTypeData.UniqueFontName;
			if (Subset) FontFileName = "1" + FontFileName; else FontFileName = "0" + FontFileName; //Fallback fonts are always subsetted. So we might have a non subsetted font and a subsetted font in the same file.

			if (TryGetValue(FontFileName, out Result)) return Result;

            Result = new TPdfEmbeddedFont(Subset, TrueTypeData);		

			this[FontFileName] = Result;
			return Result;
		}


        private static byte[] LoadFont(Font aFont, TFontEvents FontEvents)
        {
            byte[] FFontData = null;

            if (FontEvents.OnGetFontData != null)
            {
                GetFontDataEventArgs fe = new GetFontDataEventArgs(aFont);
                FontEvents.OnGetFontData(FontEvents.Sender, fe);
                if (fe.Applied)
                    FFontData = fe.FontData;
            }

#if (FRAMEWORK30 && !WPF && !__MonoCS__ && COMMENTED_OUT_BECAUSE_IT_CAN_FAIL_IN_MULTITHREAD_CASES)
            if (FFontData == null && !WithoutWPFSupport)
            {
                try
                {
                    FFontData = GetFontWithWPF(aFont, FFontData);
                }
                catch (MissingMethodException) //happens when running in mono the assemmbly compiled with VS.
                {
                    WithoutWPFSupport = true;
                    FFontData = null;
                }
                catch (FileNotFoundException)
                {
                    WithoutWPFSupport = true;
                    FFontData = null;
                }

            }
#endif
            if (FFontData == null)
            {
                String FontPath = null;
                try
                {
				    if (!WithoutGetFolderPathPermissions) FontPath = Environment.GetFolderPath(Environment.SpecialFolder.System);
                }
                catch (SecurityException)
                {
                    WithoutGetFolderPathPermissions = true; //to avoid repeating each time, as this exception is mostly going to happen in a server.
                    FontPath = null;
                }

                bool AppendFontFolder = true;
                if (FontEvents.OnGetFontFolder != null)
                {
                    GetFontFolderEventArgs ae = new GetFontFolderEventArgs(aFont, FontPath);
                    FontEvents.OnGetFontFolder(FontEvents.Sender, ae);
                    if (ae.Applied)
                        FontPath = ae.FontFolder;
                    AppendFontFolder = ae.AppendFontFolder;
                }

                if (FontPath != null && FontPath.Length > 0 && AppendFontFolder)
                {
                    FontPath = Path.Combine(Path.Combine(FontPath, TPdfTokens.GetString(TPdfToken.UpDir)), TPdfTokens.GetString(TPdfToken.FontFolder));
                }

                if (FontPath == null || FontPath.Length == 0)  //We are probably on MONO. let's try the default.
                {
                    //See if it is on /usr/X11R6/lib/X11/fonts/truetype
                    if (Directory.Exists(TPdfTokens.GetString(TPdfToken.LinuxFontFolder)))
                    {
                        FontPath = TPdfTokens.GetString(TPdfToken.LinuxFontFolder);
                    }
                    else  //Put it on the folder where the dll is, /Fonts. You could make a symbolic link from there to the place where files actually are.
                    {
                        FontPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), TPdfTokens.GetString(TPdfToken.FontFolder));
                    }
                }

                if (FontPath != null && FontPath.Length > 0 && FontPath[FontPath.Length - 1] != Path.DirectorySeparatorChar)
                    FontPath += Path.DirectorySeparatorChar;

                FFontData = PdfFontFactory.GetFontData(aFont, FontPath);
            }

            if (FFontData == null)
                PdfMessages.ThrowException(PdfErr.ErrFontNotFound, aFont.Name);

            return FFontData;
        }


#if (FRAMEWORK30 && !WPF && !__MonoCS__ && COMMENTED_OUT_BECAUSE_IT_CAN_FAIL_IN_MULTITHREAD_CASES)

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static byte[] GetFontWithWPF(Font aFont, byte[] FFontData)
        {
            System.Windows.Media.Typeface fm = new System.Windows.Media.Typeface(GetWpfFontFamily(aFont), GetWpfStyle(aFont), GetWpfFontWeight(aFont), GetWpfFontStretch(aFont));
            System.Windows.Media.GlyphTypeface gt;
            if (fm.TryGetGlyphTypeface(out gt))
            {
                using (Stream st = gt.GetFontStream())
                {
                    FFontData = new byte[st.Length];
                    st.Read(FFontData, 0, FFontData.Length);
                }
            }
            return FFontData;
        }


        private static System.Windows.Media.FontFamily GetWpfFontFamily(Font aFont)
        {
            return new System.Windows.Media.FontFamily(aFont.Name);
        }

        private static System.Windows.FontStyle GetWpfStyle(Font aFont)
        {
            if (aFont.Italic) return System.Windows.FontStyles.Italic;
            return System.Windows.FontStyles.Normal;
        }

        private static System.Windows.FontWeight GetWpfFontWeight(Font aFont)
        {
            if (aFont.Bold) return System.Windows.FontWeights.Bold;
            return System.Windows.FontWeights.Normal;
        }

        private static System.Windows.FontStretch GetWpfFontStretch(Font aFont)
        {
            return System.Windows.FontStretches.Normal;
        }


#endif
    }

	#region Base Font
	internal abstract class TPdfFont  
	{
		protected int Id;
		protected string FFontName;
		protected FontStyle FFontStyle;
		protected bool FUseKerning;

		protected int FontObjId;

		protected TPdfFont FFallbackFont;
		protected TPdfResources Resources;

		internal bool UsedInDoc;

		public string FontName {get{return FFontName;}}

		private static bool IsStandardFont(String aFontName)
		{
			aFontName = aFontName.Replace(" ",String.Empty);
			return
				String.Equals(aFontName, TPdfTokens.GetString(TPdfToken.StArial), StringComparison.CurrentCultureIgnoreCase) ||
				String.Equals(aFontName, TPdfTokens.GetString(TPdfToken.StCourier), StringComparison.CurrentCultureIgnoreCase) ||
				String.Equals(aFontName, TPdfTokens.GetString(TPdfToken.StCourierNew), StringComparison.CurrentCultureIgnoreCase) ||
				String.Equals(aFontName, TPdfTokens.GetString(TPdfToken.StMicrosoftSansSerif), StringComparison.CurrentCultureIgnoreCase) ||
				String.Equals(aFontName, TPdfTokens.GetString(TPdfToken.StMicrosoftSerif), StringComparison.CurrentCultureIgnoreCase) ||
				String.Equals(aFontName, TPdfTokens.GetString(TPdfToken.StTimesNewRoman), StringComparison.CurrentCultureIgnoreCase);
		}

		protected TPdfFont FallbackFont(int FallbackLevel)
		{
			if (FallbackLevel > 0) return Resources.CreateFallbackFont(FallbackLevel, FFontStyle, FUseKerning);
			if (FFallbackFont == null) FFallbackFont = Resources.CreateFallbackFont(FallbackLevel, FFontStyle, FUseKerning);
		    return FFallbackFont;
	    }

		private static string GetAdditionalStyleName(FontStyle AdditionalStyle)
		{
			if ((AdditionalStyle & FontStyle.Bold) != 0)
			{
				if ((AdditionalStyle & FontStyle.Italic) != 0) return "BoldItalic";
				return "Bold";
			}
			if ((AdditionalStyle & FontStyle.Italic) != 0) return "Italic";
            return "Regular";
		}

		internal static TPdfFont CreateInstance(TFontMapping Mapping, Font aFont, bool aUnicode, int aId, TFontEmbed aEmbed, TFontSubset aSubset, bool aCompress, bool aUseKerning,
			TFontEvents FontEvents, TPdfEmbeddedFontList EmbeddedFontList, TPdfResources aResources)
		{
			if (!aUnicode && 
				(Mapping == TFontMapping.ReplaceAllFonts || (Mapping == TFontMapping.ReplaceStandardFonts && IsStandardFont(aFont.Name)))
				)
				return new TPdfInternalFont(aFont, aId, aUseKerning, aResources);

			FontStyle AdditionalStyle;
			TPdfEmbeddedFont EmbeddedData = EmbeddedFontList.Add(aFont, FontEvents, aSubset == TFontSubset.Subset, aUseKerning, out AdditionalStyle);

			if (AdditionalStyle != FontStyle.Regular)
			{
				if (FlexCelTrace.HasListeners) 
					FlexCelTrace.Write(new TPdfFauxBoldOrItalicsError(FlxMessages.GetString(FlxErr.ErrFauxBoldOrItalic, aFont.Name, GetAdditionalStyleName(AdditionalStyle)), aFont.Name, AdditionalStyle));
			}

			if (aUnicode)
				return new TPdfUnicodeTrueTypeFont(aFont, aId, aSubset, aCompress, aUseKerning, EmbeddedData, AdditionalStyle, aResources);
			else
			{
				bool DoEmbed = aEmbed ==TFontEmbed.Embed || EmbeddedData.TrueTypeData.NeedsEmbed(aEmbed);
	
				if (FontEvents.OnFontEmbed!=null)
				{
					FontEmbedEventArgs ae = new FontEmbedEventArgs(aFont, DoEmbed);
					FontEvents.OnFontEmbed(FontEvents.Sender, ae);
					DoEmbed = ae.Embed;
				}

				return new TPdfWinAnsiTrueTypeFont(aFont, aId, DoEmbed, aSubset, aCompress, aUseKerning, EmbeddedData, AdditionalStyle, aResources);
			}
		}

		protected TPdfFont(Font aFont, int aId, bool aUseKerning, TPdfResources aResources)
		{
			FFontName = aFont.Name;
			FUseKerning = aUseKerning;

			FFontStyle = aFont.Style & ~FontStyle.Underline;

			Resources = aResources;

			Id = aId;
			UsedInDoc = false;
		}

		internal abstract byte[] EncodeString(string s);

		public abstract TPdfFont Fallback(char c, int FallbackLevel);
		public abstract real MeasureString(string s);
		public abstract TKernedString[] KernString(string s);
		public abstract real LineGap();
		public abstract real Ascent();
		public abstract real Descent();
		public abstract real UnderlinePosition();

		public virtual void WriteFont(TPdfStream DataStream, TXRefSection XRef)
		{
			TPdfBaseRecord.Write(DataStream, TPdfTokens.GetString(TPdfToken.FontPrefix) + Id.ToString(CultureInfo.InvariantCulture) + " ");
			FontObjId = XRef.GetNewObject(DataStream);
			TIndirectRecord.CallObj(DataStream, FontObjId);
		}
        
		public abstract void WriteFontObject(TPdfStream DataStream, TXRefSection XRef);

		public void Select(TPdfStream DataStream, real FontSize, ref string LastFont)
		{
			UsedInDoc = true;
			string NewFont = TPdfTokens.GetString(TPdfToken.FontPrefix) + Id.ToString(CultureInfo.InvariantCulture) + " "+
				PdfConv.CoordsToString(FontSize) + " " + TPdfTokens.GetString(TPdfToken.CommandFont);

			if (LastFont == NewFont) return;
			LastFont = NewFont;
			TPdfBaseRecord.WriteLine(DataStream, NewFont);
		}

		public abstract void AddString(string s);
		public abstract bool AddChar(int code, int FallbackLevel);
	}
	#endregion

	#region Base True Type
	internal abstract class TPdfBaseTrueTypeFont: TPdfFont
	{
		#region Protected
		protected TPdfTrueType FTrueTypeData;
		protected TPdfEmbeddedFont EmbeddedData;
		protected bool FCompress;
		protected FontStyle AdditionalStyle;
		#endregion

		protected TPdfBaseTrueTypeFont(Font aFont, int aId, bool aCompress, bool aUseKerning, TPdfEmbeddedFont aEmbeddedData, FontStyle aAdditionalStyle, TPdfResources aResources):
			base(aFont, aId, aUseKerning, aResources)
		{
			FCompress = aCompress;
			EmbeddedData = aEmbeddedData;
			FTrueTypeData = aEmbeddedData.TrueTypeData;
			AdditionalStyle = aAdditionalStyle;
		}

		protected string GetFontName(bool Embed, bool Subset)
		{
			string Result = "/";
			if (Subset && Embed) Result +="ABGVAA+";
			Result +=FTrueTypeData.PostcriptName;
			if ((AdditionalStyle & FontStyle.Italic)!=0)
			{
				if ((AdditionalStyle & FontStyle.Bold)!=0)
					Result += TPdfTokens.GetString(TPdfToken.BoldItalic);
				else
					Result+=TPdfTokens.GetString(TPdfToken.Italic);
			}
			else
				if ((AdditionalStyle & FontStyle.Bold) !=0)
				Result += TPdfTokens.GetString(TPdfToken.Bold);

			return Result;

		}

		public override real LineGap()
		{
			return FTrueTypeData.LineGap*1000F/FTrueTypeData.UnitsPerEm;
		}

		public override real Ascent()
		{
			return FTrueTypeData.Ascent*1000F/FTrueTypeData.UnitsPerEm;
		}

		public override real Descent()
		{
			return FTrueTypeData.Descent*1000F/FTrueTypeData.UnitsPerEm;
		}

		public override real UnderlinePosition()
		{
			return FTrueTypeData.UnderlinePosition*1000F / FTrueTypeData.UnitsPerEm;
		}
      
		protected void SaveFontDescriptor(TPdfStream DataStream, int FontDescriptorId, TXRefSection XRef, bool Embed, bool Subset)
		{

			XRef.SetObjectOffset(FontDescriptorId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, FontDescriptorId);
			TDictionaryRecord.BeginDictionary(DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.FontDescriptorName));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.FontNameName, GetFontName(Embed, Subset));

			string BBox = 
				TPdfTokens.GetString(TPdfToken.OpenArray) +
				FTrueTypeData.BoundingBox.GetBBox(FTrueTypeData.UnitsPerEm)+
				TPdfTokens.GetString(TPdfToken.CloseArray);

			TDictionaryRecord.SaveKey(DataStream, TPdfToken.FontBBoxName, BBox);

			TDictionaryRecord.SaveKey(DataStream, TPdfToken.FontAscentName, (int)Math.Round(FTrueTypeData.Ascent*1000F/FTrueTypeData.UnitsPerEm));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.FontDescentName, (int)Math.Round(FTrueTypeData.Descent*1000F/FTrueTypeData.UnitsPerEm));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.CapHeightName, (int)Math.Round(FTrueTypeData.CapHeight*1000F/FTrueTypeData.UnitsPerEm));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.ItalicAngleName, FTrueTypeData.ItalicAngle.ToString(CultureInfo.InvariantCulture));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.StemVName, 80); //StemV info is not present on a true type font. 80 seems to be what adobe writes.
			
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.FlagsName, FTrueTypeData.FontFlags);

			if (Embed)
			{
				TDictionaryRecord.SaveKey(DataStream, TPdfToken.FontFile2Name, TIndirectRecord.GetCallObj(EmbeddedData.GetStreamId(DataStream, XRef)));
			}


			TDictionaryRecord.EndDictionary(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);

			EmbeddedData.WriteFont(DataStream, XRef, FCompress);
		}

		public override TPdfFont Fallback(char c, int FallbackLevel)
		{
			if (FTrueTypeData.HasGlyph((int)c)) return this;
			TPdfFont Fallbk = FallbackFont(FallbackLevel);
			if (Fallbk != null) return Fallbk.Fallback(c, FallbackLevel + 1);
			return null;
		}

	}

	#endregion

	#region WinAnsi True Type Font
	internal class TPdfWinAnsiTrueTypeFont: TPdfBaseTrueTypeFont
	{
		private bool Embed;
		private bool Subset;
		protected int FirstChar;
		protected int LastChar;

		private int FontDescriptorId;

		public TPdfWinAnsiTrueTypeFont(Font aFont, int aId, bool aEmbed, TFontSubset aSubset, bool aCompress, bool aUseKerning,
			TPdfEmbeddedFont aEmbeddedData, FontStyle aAdditionalStyle, TPdfResources aResources): base(aFont, aId, aCompress, aUseKerning, aEmbeddedData, aAdditionalStyle, aResources)
		{		
			Embed = aEmbed;
			Subset = aSubset == TFontSubset.Subset;
			FirstChar=-1;
			LastChar=0;
		}

		private static string EncodingType()
		{
			return TPdfTokens.GetString(TPdfToken.WinAnsiEncodingName);
		}

		private static string FontType()
		{
			return TPdfTokens.GetString(TPdfToken.TrueTypeName);
		}

		private void SaveWidths(TPdfStream DataStream)
		{
			//Docs say it is prefered to save as indirect object, but acrobat writes it directly.
			//Widths object.
			TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.WidthsName)+" "+TPdfTokens.GetString(TPdfToken.OpenArray));
			int fc=FirstChar;if (fc<0) fc=0;
            
			for (int i=fc; i<=LastChar; i++) //characters on winansi are NOT the same as in low byte unicode.  For example char 0x92 (146) is a typographic ' in winansi, not defined in unicode.
			{
				int z = (int)CharUtils.GetUniFromWin1252_PDF((byte)i);
				TPdfBaseRecord.WriteLine(DataStream, PdfConv.CoordsToString(Math.Round(FTrueTypeData.GlyphWidth(FTrueTypeData.Glyph(z, false))))); //don't log the erorr here, character doesn't need to exist.
			}

			TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CloseArray));
		}

		internal override byte[] EncodeString(string s)
		{
			return CharUtils.GetWin1252Bytes_PDF(s);
		}

		public override real MeasureString(string s)
		{
			real Result = 0;
			//Here we don't have chars > 255, so there are no unicode surrogates.
			int[] bi = new int[s.Length];
			bool[] ignore = new bool[s.Length];
			for (int i =0; i<s.Length;i++)
			{
				TPdfFont f = Fallback(s[i], 0);
				if (f != this && f!= null) 
				{
					Result += f.MeasureString(new String(s[i], 1));
					ignore[i] = true;
					bi[i] = 0;
				}
				else bi[i] = FTrueTypeData.Glyph(s[i], true); //This returns the Glyph for an unicode character, so s[i] is ok. (s is unicode)
			}

			return Result + FTrueTypeData.MeasureString(bi, ignore);
		}

		public override TKernedString[] KernString(string s)
		{
			//Here we don't have chars > 255, so there are no unicode surrogates.
			int[] bi = new int[s.Length];
			for (int i =0; i<s.Length;i++)
				bi[i] = FTrueTypeData.Glyph(s[i], false);

			return FTrueTypeData.KernString(s, bi);
		}


		public override void WriteFont(TPdfStream DataStream, TXRefSection XRef)
		{
			base.WriteFont (DataStream, XRef);
			FontDescriptorId = XRef.GetNewObject(DataStream); //Font descriptor.
		}

		public override void WriteFontObject(TPdfStream DataStream, TXRefSection XRef)
		{
			XRef.SetObjectOffset(FontObjId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, FontObjId);
			TDictionaryRecord.BeginDictionary(DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.FontName));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.SubtypeName, FontType());
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.BaseFontName, GetFontName(Embed, Subset));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.EncodingName, EncodingType());
			long fc=FirstChar;if (fc<0) fc=0;
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.FirstCharName, PdfConv.LongToString(fc));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.LastCharName, PdfConv.LongToString(LastChar));

			XRef.SetObjectOffset(FontDescriptorId, DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.FontDescriptorName, TIndirectRecord.GetCallObj(FontDescriptorId));
            
			SaveWidths(DataStream);
 
			TDictionaryRecord.EndDictionary(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);
            
			SaveFontDescriptor(DataStream, FontDescriptorId, XRef, Embed, Subset);
		}


		public override bool AddChar(int code, int FallbackLevel)
		{
			bool Result = FTrueTypeData.HasGlyph(code);

			if (!Result)
			{
				TPdfFont Fallbk = FallbackFont(FallbackLevel);
				if (Fallbk != null)
				{
					if (Fallbk.AddChar(code, FallbackLevel + 1)) return true;
				}
			}

			if (Result || FallbackLevel == 0)
			{
				//characters on winansi are NOT the same than low byte unicode.f.i. char 0x92 is not defined in unicode.
				byte acode = CharUtils.GetWin1252Bytes_PDF(code);

				if(FirstChar<0 || acode< FirstChar)
					FirstChar=acode;
				if(acode> LastChar)
					LastChar=acode;

				if (Embed && Subset)EmbeddedData.AddGlyphFromChar(code);
			}
			
			return Result;
		}


		public override void AddString(string s)
		{
			//characters on winansi are NOT the same than low byte unicode.f.i. char 0x92 is not defined in unicode.
			for (int i=0; i< s.Length; i++)
			{         
				int code= s[i];
				AddChar(code, 0);
			}
		}

	}
	#endregion

	#region UNICODE True Type Font
	internal class TPdfUnicodeTrueTypeFont: TPdfBaseTrueTypeFont
	{
		TToUnicode ToUnicodeData;
		private int FontDescriptorId;
		private int CIDFontId;
		private int ToUnicodeId;
		private int ToUnicodeLenId;
		private TUsedCharList UsedChars;
		bool Subset;

		public TPdfUnicodeTrueTypeFont(Font aFont, int aId, TFontSubset aSubset, bool aCompress, bool aUseKerning,
			TPdfEmbeddedFont aEmbeddedData, FontStyle aAdditionalStyle, TPdfResources aResources): base(aFont, aId, aCompress, aUseKerning, aEmbeddedData, aAdditionalStyle, aResources)
		{
			UsedChars = new TUsedCharList();
			ToUnicodeData = new TToUnicode();
			Subset = aSubset == TFontSubset.Subset;
		}

		private static string EncodingType()
		{
			return TPdfTokens.GetString(TPdfToken.IdentityHName);
		}

		internal static void MotorolaSetWord(byte[] Data, int tPos, int number) 
		{  
			unchecked
			{
				Data[tPos+1]=(byte) number;
				Data[tPos]=(byte) (number >> 8);
			}
		}

		internal override byte[] EncodeString(string s)
		{
			byte[] Result = new byte[2*s.Length];
			for (int i=0; i<s.Length;i++)
			{
				int gl = FTrueTypeData.Glyph((int) s[i], true);
				MotorolaSetWord(Result,i*2,  EmbeddedData.GetNewGlyphFromOldGlyph(gl));
			}

			return Result;
		}

		public override real MeasureString(string s)
		{
			real Result = 0;
			int[] b = new int[s.Length];
			bool[] ignore = new bool[s.Length];
			for (int i=0; i<b.Length;i++) 
			{
				TPdfFont f = Fallback(s[i], 0);
				if (f != this && f!= null) 
				{
					Result += f.MeasureString(new String(s[i], 1));
					ignore[i] = true;
					b[i] = 0;
				}
				else b[i]=FTrueTypeData.Glyph(s[i], true);
			}
			return Result + FTrueTypeData.MeasureString(b, ignore);
		}

		public override TKernedString[] KernString(string s)
		{
			int[] b = new int[s.Length];
			for (int i=0; i<b.Length;i++) b[i]=FTrueTypeData.Glyph(s[i], false);
			return FTrueTypeData.KernString(s, b);
		}

		public override void WriteFont(TPdfStream DataStream, TXRefSection XRef)
		{
			base.WriteFont (DataStream, XRef);
			FontDescriptorId = XRef.GetNewObject(DataStream); //Font descriptor.
			CIDFontId = XRef.GetNewObject(DataStream); //CID Font.
			ToUnicodeId = XRef.GetNewObject(DataStream); //ToUnicode
			ToUnicodeLenId = XRef.GetNewObject(DataStream); //ToUnicodeLength
		}

		private void SaveWidths(TPdfStream DataStream)
		{
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.DWName, UsedChars.DW());
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.WName, UsedChars.W());			
		}

		//For unicode fonts, we actually need to write 2 fonts, one type 0 and one CID.
		public override void WriteFontObject(TPdfStream DataStream, TXRefSection XRef)
		{
			//Save Type 0 Font.
			XRef.SetObjectOffset(FontObjId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, FontObjId);
			TDictionaryRecord.BeginDictionary(DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.FontName));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.SubtypeName, TPdfTokens.GetString(TPdfToken.Type0Name));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.BaseFontName, GetFontName(true, Subset));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.EncodingName, EncodingType());
 
			XRef.SetObjectOffset(CIDFontId, DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.DescendantFontsName,
				TPdfTokens.GetString(TPdfToken.OpenArray) + TIndirectRecord.GetCallObj(CIDFontId) + TPdfTokens.GetString(TPdfToken.CloseArray));

			TDictionaryRecord.SaveKey(DataStream, TPdfToken.ToUnicodeName, TIndirectRecord.GetCallObj(ToUnicodeId));

			TDictionaryRecord.EndDictionary(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);

			//Save CID Font.
			XRef.SetObjectOffset(CIDFontId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, CIDFontId);
			TDictionaryRecord.BeginDictionary(DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.FontName));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.SubtypeName, TPdfTokens.GetString(TPdfToken.CIDFontType2Name));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.BaseFontName, GetFontName(true, Subset));
			TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CIDSystemInfo));
			TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CIDToGIDMap));

			SaveWidths(DataStream);

			XRef.SetObjectOffset(FontDescriptorId, DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.FontDescriptorName, TIndirectRecord.GetCallObj(FontDescriptorId));

			TDictionaryRecord.EndDictionary(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);

			SaveToUnicode(DataStream, XRef);
			SaveFontDescriptor(DataStream, FontDescriptorId, XRef, true, Subset);
		}

		private void SaveToUnicode(TPdfStream DataStream, TXRefSection XRef)
		{
			XRef.SetObjectOffset(ToUnicodeId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, ToUnicodeId);

			XRef.SetObjectOffset(ToUnicodeLenId, DataStream);
			TStreamRecord.BeginSave(DataStream, ToUnicodeLenId, FCompress);
			long StartStream = DataStream.Position;
			bool Compressing = DataStream.Compress;
			try
			{
				DataStream.Compress = FCompress;
				WriteToUnicodeData(DataStream);
			}
			finally
			{
				DataStream.Compress = Compressing;
			}
			long EndStream = DataStream.Position;
			TStreamRecord.EndSave(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);

			XRef.SetObjectOffset(ToUnicodeLenId, DataStream);

			TIndirectRecord.SaveHeader(DataStream, ToUnicodeLenId);
			TPdfBaseRecord.WriteLine(DataStream, (EndStream - StartStream).ToString(CultureInfo.InvariantCulture));
			TIndirectRecord.SaveTrailer(DataStream);
		}

		private void WriteToUnicodeData(TPdfStream DataStream)
		{
			string s = TPdfTokens.GetString(TPdfToken.ToUnicodeData).Replace("\r","");
			TPdfBaseRecord.WriteLine(DataStream, s);

			TPdfBaseRecord.WriteLine(DataStream, ToUnicodeData.GetData());

			s = TPdfTokens.GetString(TPdfToken.ToUnicodeData2).Replace("\r","");
			TPdfBaseRecord.WriteLine(DataStream, s);
		}


		public override bool AddChar(int code, int FallbackLevel)
		{
			bool Result = FTrueTypeData.HasGlyph(code);
			if (!Result)
			{
				TPdfFont Fallbk = FallbackFont(FallbackLevel);
				if (Fallbk != null)
				{
					if (Fallbk.AddChar(code, FallbackLevel + 1)) return true;
				}
			}

			if (Result || FallbackLevel == 0)
			{
				int ccode = FTrueTypeData.Glyph(code, true);
				int newcode = EmbeddedData.GetNewGlyphFromOldGlyph(ccode);
				UsedChars.Add((char)newcode, FTrueTypeData.GlyphWidth(ccode));
				ToUnicodeData.Add(newcode, (int)code);
			}

			return Result;
		}

		public override void AddString(string s)
		{
			for (int i=0; i< s.Length; i++)
			{         
				AddChar((int)s[i], 0);
			}
		}
	}

	#endregion

	#region Internal font

	internal enum TFontType
	{
		Times = 0,
		Helvetica = 1,
		Symbol = 2,
		Courier = 3
	}

	internal class TPdfInternalFont: TPdfFont  
	{
		private TFontType FFontType;
		private int[] Width;
		private TKerningTable Kern;

		private static readonly int[] FAscent = {683, 718, 693, 629};
		private static readonly int[] FDescent = {217, 207, 216, 157};
		private static readonly int[] FUnderlinePosition = {-100, -100, -100, -100};
		private static readonly int[] FLineGap = {150,150,150,0};


		public TPdfInternalFont(Font aFont, int aId, bool aUseKerning, TPdfResources aResources): base(aFont, aId, aUseKerning, aResources)
		{
			//LOGFONT lf = new LOGFONT();  This does not work, and it is unmanaged also!

			string FontName = ";" + aFont.Name.Replace(" ",String.Empty).ToUpper(CultureInfo.InvariantCulture)+";";

			if (TPdfTokens.GetString(TPdfToken.StSymbolFonts).IndexOf(FontName)>=0)
			{
				FFontName = TPdfTokens.GetString(TPdfToken.PsSymbol);
				FFontType = TFontType.Symbol;
			}
			else
				if (TPdfTokens.GetString(TPdfToken.StSerifFonts).IndexOf(FontName)>=0)
			{
				FFontName = TPdfTokens.GetString(TPdfToken.PsTimes);
				FFontType = TFontType.Times;
			}
			else
				if (TPdfTokens.GetString(TPdfToken.StFixedFonts).IndexOf(FontName)>=0)
			{
				FFontName = TPdfTokens.GetString(TPdfToken.PsCourier);
				FFontType = TFontType.Courier;
			}
			else
			{
				FFontName = TPdfTokens.GetString(TPdfToken.PsHelvetica);
				FFontType = TFontType.Helvetica;
			}

			Width = TInternalFontMetrics.Width(FFontType, FFontStyle);
			if (FUseKerning) Kern = TInternalFontMetrics.Kern(FFontType, FFontStyle);
            
		}

		internal override byte[] EncodeString(string s)
		{
			return CharUtils.GetWin1252Bytes_PDF(s);
		}

		private string GetFontName()
		{
			string Result = "/"+FFontName;

			//Times new roman fonts are differently named. We have italic instead of oblique and -Roman for normal.
			if (String.Compare(Result, 1, TPdfTokens.GetString(TPdfToken.PsTimes), 0, Result.Length, StringComparison.CurrentCulture) == 0)
			{
				if ((FFontStyle & FontStyle.Italic)  !=0)
				{
					if ((FFontStyle & FontStyle.Bold) !=0)
						Result+=TPdfTokens.GetString(TPdfToken.PsBoldItalic);
					else
						Result+=TPdfTokens.GetString(TPdfToken.PsItalic);
				}
				else
					if ((FFontStyle & FontStyle.Bold) !=0)
					Result += TPdfTokens.GetString(TPdfToken.PsBold);

				else Result +=TPdfTokens.GetString(TPdfToken.PsRoman);
			}
            
			else  //Not times.
			{
				if ((FFontStyle & FontStyle.Italic)  !=0)
				{
					if ((FFontStyle & FontStyle.Bold) !=0)
						Result+=TPdfTokens.GetString(TPdfToken.PsBoldOblique);
					else
						Result+=TPdfTokens.GetString(TPdfToken.PsOblique);
				}
				else
					if ((FFontStyle & FontStyle.Bold) !=0)
					Result += TPdfTokens.GetString(TPdfToken.PsBold);
            
			}
			return Result;
		}

		private static string EncodingType()
		{
			return TPdfTokens.GetString(TPdfToken.WinAnsiEncodingName);
		}

		public override TPdfFont Fallback(char c, int FallbackLevel)
		{
			if (CharUtils.IsWin1252((int)c)) return this;
			
			TPdfFont Fallbk = FallbackFont(FallbackLevel);
			if (Fallbk != null) return Fallbk.Fallback(c, FallbackLevel + 1);
			return null;
		}

		public override TKernedString[] KernString(string s)
		{
			//Here we don't have chars > 255, so there are no unicode surrogates.
			//Kern and Width tables are for Win 1252, so we need to pass that byte in the si[] array.
			byte[] b = CharUtils.GetWin1252Bytes_PDF(s);
			int[] si = new int[b.Length];
			for (int i=0;i<s.Length;i++) si[i]=(int)b[i];
			return FontMeasures.KernString(s, si, Kern, 1000);
		}
 

		public override real MeasureString(string s)
		{
			//Here we don't have chars > 255, so there are no unicode surrogates.
			//Kern and Width tables are for Win 1252, so we need to pass that byte in the si[] array.
			byte[] b = CharUtils.GetWin1252Bytes_PDF(s);
			int[] si = new int[b.Length];
			bool[] ignore = new bool[b.Length];
			for (int i=0;i<s.Length;i++) si[i]=(int)b[i];
			return FontMeasures.MeasureString(si, Width, Kern, 1000, ignore);
		}

		public override real LineGap()
		{
			return FLineGap[(int)FFontType];
		}
 
		public override real Ascent()
		{
			return FAscent[(int)FFontType];
		}

		public override real Descent()
		{
			return FDescent[(int)FFontType];
		}

		public override real UnderlinePosition()
		{
			return FUnderlinePosition[(int)FFontType];
		}


		public override void WriteFontObject(TPdfStream DataStream, TXRefSection XRef)
		{
			XRef.SetObjectOffset(FontObjId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, FontObjId);
			TDictionaryRecord.BeginDictionary(DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.FontName));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.SubtypeName, TPdfTokens.GetString(TPdfToken.Type1Name));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.BaseFontName, GetFontName());
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.EncodingName,EncodingType());

			TDictionaryRecord.EndDictionary(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);
		}

		public override bool AddChar(int code, int FallbackLevel)
		{
			return false;
		}

		public override void AddString(string s)
		{
			//Nothing here.
		}
	}
	#endregion
}



