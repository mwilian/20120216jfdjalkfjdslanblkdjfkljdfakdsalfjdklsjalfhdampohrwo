#region Using directives
using System;
using System.Text;
using System.IO;
using System.Globalization;
using System.Collections.Generic;

using FlexCel.Core;

#if (WPF)
using RectangleF = System.Windows.Rect;
using PointF = System.Windows.Point;
using ColorBlend = System.Windows.Media.GradientStopCollection;
using real = System.Double;
using System.Windows.Media;
using System.Windows;
#else
using real = System.Single;
using System.Drawing;
using System.Drawing.Drawing2D;
#endif


#endregion

namespace FlexCel.Pdf
{
    internal class TPdfResources
    {
        private TPdfFontList Fonts;
		private TPdfEmbeddedFontList EmbeddedFontList;  //While Fonts is different for Arial and arial italic (or arial with unicode chars), EmbeddedFontList is the same, because the font has to be embedded just once.
        private List<TPdfImage> Images;
		private List<TPdfHatch> HatchPatterns;
		private List<TPdfImageTexture> ImageTexturePatterns;
		private List<TPdfGradient> GradientPatterns;
        private List<TPdfFunction> Functions;
        private List<TPdfTransparency> GStates;
        private int PatternColorSpaceId;

		private TFontEvents FFontEvents;
		private bool FCompress;
		private string[] FFallbackFontList;

        internal TPdfResources(string aFallbackFontList, bool aCompress, TFontEvents FontEvents)
        {
            Fonts = new TPdfFontList();
			EmbeddedFontList = new TPdfEmbeddedFontList();
            Images = new List<TPdfImage>();
			HatchPatterns = new List<TPdfHatch>();
			ImageTexturePatterns = new List<TPdfImageTexture>();
			GradientPatterns = new List<TPdfGradient>();
            Functions = new List<TPdfFunction>();
            GStates = new List<TPdfTransparency>();

			FCompress = aCompress;
			FFontEvents = FontEvents;
			if (aFallbackFontList != null) FFallbackFontList = aFallbackFontList.Split(';');
		}

		public TPdfFont CreateFallbackFont(int FallbackLevel, FontStyle aFontStyle, bool aUseKerning)
		{
			if (FFallbackFontList == null || FFallbackFontList.Length <= FallbackLevel) return null; //no fallback.

#if (!WPF)
			using (
#endif
            Font aFont = new Font(FFallbackFontList[FallbackLevel], 10, aFontStyle)
#if (WPF)
;
#else
			)
#endif
			{
				TPdfFont Result = GetFont(TFontMapping.DontReplaceFonts, aFont, true, TFontEmbed.Embed, TFontSubset.Subset, aUseKerning);
				return Result;
			}
		}
		

        internal void SaveResourceDesc(TPdfStream DataStream, TXRefSection XRef, bool IncludeDict)
        {
            if (IncludeDict) TDictionaryRecord.BeginDictionary(DataStream);
            SaveFonts(DataStream, XRef);
            SaveImages(DataStream, XRef);
            SavePatterns(DataStream, XRef);
            SaveGStates(DataStream, XRef);
            if (IncludeDict) TDictionaryRecord.EndDictionary(DataStream);
        }

        private void SaveFonts(TPdfStream DataStream, TXRefSection XRef)
        {
            int aCount = Fonts.Count;
            if (aCount <= 0) return;
            TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.FontName));
            TDictionaryRecord.BeginDictionary(DataStream);

            Fonts.CreateOrderedKeys();
            foreach (string s in Fonts.Keys)
            {
                Fonts.OrderedKeys.Add(s);
            }

            Fonts.OrderedKeys.Sort(); //to keep the reference pdfs always the same.

            foreach (string s in Fonts.OrderedKeys)
            {
                TPdfFont pf = ((TPdfFont)Fonts[s]);
				if (pf.UsedInDoc) pf.WriteFont(DataStream, XRef);
            }
            TDictionaryRecord.EndDictionary(DataStream);
        }

        private void SaveImages(TPdfStream DataStream, TXRefSection XRef)
        {
            int aCount = Images.Count;
            if (aCount <= 0) return;
            TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.XObjectName));
            TDictionaryRecord.BeginDictionary(DataStream);
            for (int i = 0; i < aCount; i++)
            {
                Images[i].WriteImage(DataStream, XRef);
            }
            TDictionaryRecord.EndDictionary(DataStream);
        }

        private void SavePatterns(TPdfStream DataStream, TXRefSection XRef)
        {
			int HatchCount = HatchPatterns.Count;
			int ImageTextureCount = ImageTexturePatterns.Count;
			int GradientCount = GradientPatterns.Count;
            if (HatchCount + GradientCount + ImageTextureCount <= 0) return;
            TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.PatternName));
			TDictionaryRecord.BeginDictionary(DataStream);
			for (int i = 0; i < HatchCount; i++)
			{
				HatchPatterns[i].WritePattern(DataStream, XRef);
			}

			for (int i = 0; i < ImageTextureCount; i++)
			{
				ImageTexturePatterns[i].WritePattern(DataStream, XRef);
			}
            
            for (int i = 0; i < GradientCount; i++)
            {
                GradientPatterns[i].WritePattern(DataStream, XRef);
            }

            TDictionaryRecord.EndDictionary(DataStream);
            if (HatchCount <= 0) return;

            //Pattern color space
            TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.ColorSpaceName));
            TDictionaryRecord.BeginDictionary(DataStream);
            PatternColorSpaceId = TPdfHatch.WriteColorSpace(DataStream, XRef);
            TDictionaryRecord.EndDictionary(DataStream);
        }

        private void SaveGStates(TPdfStream DataStream, TXRefSection XRef)
        {
            int aCount = GStates.Count;
            if (aCount <= 0) return;
            TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.ExtGStateName));
            TDictionaryRecord.BeginDictionary(DataStream);
            for (int i = 0; i < aCount; i++)
            {
                GStates[i].WriteGState(DataStream, XRef);
            }
            TDictionaryRecord.EndDictionary(DataStream);
        }

        internal void SaveObjects(TPdfStream DataStream, TXRefSection XRef)
        {
            if (Fonts.Count > 0)
            {
                foreach (string s in Fonts.OrderedKeys)
                {
					TPdfFont pf = ((TPdfFont)Fonts[s]);
					if (pf.UsedInDoc) pf.WriteFontObject(DataStream, XRef);
                }
            }

            int aCount = Images.Count;
            for (int i = 0; i < aCount; i++)
            {
                Images[i].WriteImageObject(DataStream, XRef);
            }

            aCount = HatchPatterns.Count;
            for (int i = 0; i < aCount; i++)
            {
                HatchPatterns[i].WritePatternObject(DataStream, XRef);
            }

            if (aCount > 0)
            {
                TPdfHatch.WriteColorSpaceObject(DataStream, XRef, PatternColorSpaceId);
            }

			aCount = ImageTexturePatterns.Count;
			for (int i = 0; i < aCount; i++)
			{
				ImageTexturePatterns[i].WritePatternObject(DataStream, XRef);
			}

            aCount = GradientPatterns.Count;
            for (int i = 0; i < aCount; i++)
            {
                GradientPatterns[i].WritePatternObject(DataStream, XRef);
            }

            aCount = Functions.Count;
            for (int i = 0; i < aCount; i++)
            {
                Functions[i].WriteFunctionObject(DataStream, XRef, FCompress);
            }

            aCount = GStates.Count;
            for (int i = 0; i < aCount; i++)
            {
                GStates[i].WriteGStateObject(DataStream, XRef);
            }
        }

        internal TPdfFont SelectFont(TPdfStream DataStream, TFontMapping Mapping, Font aFont, string s, TFontEmbed aEmbed, TFontSubset aSubset, bool aUseKerning,
            ref string LastFont)
        {
            TPdfFont Result = GetFont(Mapping, aFont, s, aEmbed, aSubset, aUseKerning);
            Result.AddString(s);
            Result.Select(DataStream, aFont.SizeInPoints, ref LastFont);
            return Result;
        }

        internal static string GetKey(Font aFont, bool NeedsUnicode)
		{
			string u = NeedsUnicode ? "1" : "0";
			string i = aFont.Italic ? "1" : "0";
			string b = aFont.Bold ? "1" : "0";
			return  u + i + b + aFont.Name;
		}

		internal TPdfFont GetFont(TFontMapping Mapping, Font aFont, string s, TFontEmbed aEmbed, TFontSubset aSubset, bool aUseKerning)
		{
			bool NeedsUnicode = CharUtils.IsWin1252(s);
			return GetFont(Mapping, aFont, NeedsUnicode, aEmbed, aSubset, aUseKerning);
		}

        internal TPdfFont GetFont(TFontMapping Mapping, Font aFont, bool NeedsUnicode, TFontEmbed aEmbed, TFontSubset aSubset, bool aUseKerning)
        {
			string key = GetKey(aFont, NeedsUnicode);
            TPdfFont SearchFont;

			if (Fonts.ContainsKey(key))
			{
				SearchFont = ((TPdfFont)Fonts[key]);
			}

			else
			{
				SearchFont = TPdfFont.CreateInstance(Mapping, aFont, NeedsUnicode, Fonts.Count, aEmbed, aSubset, FCompress, aUseKerning, FFontEvents, EmbeddedFontList, this);
				Fonts.Add(key, SearchFont);
			}

            return SearchFont;
        }

        internal void SelectImage(TPdfStream DataStream, Image aImage, Stream ImageData, long transparentColor, bool defaultToJpg)
        {
            TPdfImage SearchImage = AddImage(aImage, ImageData, transparentColor, defaultToJpg);
            SearchImage.Select(DataStream);
        }

        internal TPdfImage AddImage(Image aImage, Stream ImageData, long transparentColor, bool defaultToJpg)
        {
            TPdfImage SearchImage = new TPdfImage(aImage, Images.Count, ImageData, transparentColor, defaultToJpg);
            int Index = Images.BinarySearch(0, Images.Count, SearchImage, null);  //Only BinarySearch compatible with CF.

            if (Index < 0)
                Images.Insert(~Index, SearchImage);
            else SearchImage = Images[Index];
            return SearchImage;
        }

        internal void SelectPattern(TPdfStream DataStream, HatchStyle aStyle, Color aColor)
        {
            TPdfHatch SearchPattern = new TPdfHatch(HatchPatterns.Count, aStyle);
            int Index = HatchPatterns.BinarySearch(0, HatchPatterns.Count, SearchPattern, null);  //Only BinarySearch compatible with CF.

            if (Index < 0)
                HatchPatterns.Insert(~Index, SearchPattern);
            else SearchPattern = HatchPatterns[Index];

            SearchPattern.Select(DataStream, aColor);
        }

		internal void SelectPattern(TPdfStream DataStream, Image aImage, real[] PatternMatrix)
		{
			TPdfImageTexture SearchPattern = new TPdfImageTexture(ImageTexturePatterns.Count, aImage, PatternMatrix);
			int Index = ImageTexturePatterns.BinarySearch(0, ImageTexturePatterns.Count, SearchPattern, null);  //Only BinarySearch compatible with CF.

			if (Index < 0)
				ImageTexturePatterns.Insert(~Index, SearchPattern);
			else SearchPattern = ImageTexturePatterns[Index];

			SearchPattern.Select(DataStream);
		}

        internal void SelectGradient(TPdfStream DataStream, TGradientType aGradientType, ColorBlend aBlendColors, RectangleF aCoords, PointF aCenterPoint, RectangleF aRotatedCoords, string DrawingMatrix)
        {
            TPdfGradient SearchGradient = GetGradient(aGradientType, aBlendColors, aCoords, aCenterPoint, aRotatedCoords, DrawingMatrix);
            SearchGradient.Select(DataStream);
        }

        internal TPdfGradient GetGradient(TGradientType aGradientType, ColorBlend aBlendColors, RectangleF aCoords, PointF aCenterPoint, RectangleF RotatedCoords, string DrawingMatrix)
        {
            TPdfGradient SearchGradient = new TPdfGradient(GradientPatterns.Count, aGradientType, aBlendColors, aCoords, aCenterPoint, RotatedCoords, DrawingMatrix, Functions);
            int Index = GradientPatterns.BinarySearch(0, GradientPatterns.Count, SearchGradient, null);  //Only BinarySearch compatible with CF.

            if (Index < 0)
                GradientPatterns.Insert(~Index, SearchGradient);
            else SearchGradient = GradientPatterns[Index];

            return SearchGradient;
        }

        internal void SelectTransparency(TPdfStream DataStream, int Alpha, TPdfToken aOperator)
        {
            TPdfTransparency Transparency = GetTransparency(Alpha, aOperator, null, null);
            Transparency.Select(DataStream);
        }

        internal void SelectTransparency(TPdfStream DataStream, int Alpha, TPdfToken aOperator, string aSMask, string aBBox)
        {
            TPdfTransparency Transparency = GetTransparency(Alpha, aOperator, aSMask, aBBox);
            Transparency.Select(DataStream);
        }

        internal TPdfTransparency GetTransparency(int Alpha, TPdfToken aOperator, string aSMask, string aBBox)
        {
            TPdfTransparency SearchTransparency = new TPdfTransparency(GStates.Count, Alpha, aOperator, aSMask, aBBox);
            int Index = GStates.BinarySearch(0, GStates.Count, SearchTransparency, null);  //Only BinarySearch compatible with CF.

            if (Index < 0)
                GStates.Insert(~Index, SearchTransparency);
            else SearchTransparency = GStates[Index];

            return SearchTransparency;
        }

    }
}
