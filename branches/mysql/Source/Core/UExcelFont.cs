using System;
using System.Collections.Generic;
using System.Globalization;

#if (MONOTOUCH)
	using real = System.Single;
	using System.Drawing;
    using Color = MonoTouch.UIKit.UIColor;
    using Font = MonoTouch.UIKit.UIFont;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	using real = System.Double;
	using System.Windows;
	using System.Windows.Media;
	using System.Windows.Controls;
	#else
	using real = System.Single;
	using System.Drawing;
	#endif
#endif

namespace FlexCel.Core
{
#if (MONOTOUCH)
	/// <summary>
	/// A substitute for FontStyle when running in MonoTouch. 
	/// </summary>
	public enum FontStyle
	{
		/// <summary>
		/// Normal Font. 
		/// </summary>
		Regular,
		
		/// <summary>
		/// Bold Font. 
		/// </summary>
		Bold,
		
		/// <summary>
		/// Italic Font. 
		/// </summary>
		Italic,
		
		/// <summary>
		/// Strikeout Font. 
		/// </summary>
		Strikeout,
		
		/// <summary>
		/// Underlined font. 
		/// </summary>
		Underline
	}
#endif
#if (WPF)
	/// <summary>
	/// Utility methods to create normal fonts from Excel ones.
	/// </summary>
    public sealed class ExcelFont
    {
        /// <summary>
        /// Assigns an Excel font to a TextBlock
        /// </summary>
        /// <param name="tb">TextBlock to which assign the font.</param>
        /// <param name="Fx">Font to assign.</param>
        public static void SetFont(TextBlock tb, TFlxFont Fx)
        {
            tb.FontFamily = new FontFamily(Fx.Name);
            tb.FontSize = Fx.Size20 / 20.0 * FlxConsts.PixToPoints; //fontsize is in pixels, not points
            if ((Fx.Style & TFlxFontStyles.Italic) != 0) tb.FontStyle = FontStyles.Italic; else tb.FontStyle = FontStyles.Normal;
            if ((Fx.Style & TFlxFontStyles.Bold) != 0) tb.FontWeight = FontWeights.Bold; else tb.FontWeight = FontWeights.Normal;

#if (!SILVERLIGHT)
            if ((Fx.Style & TFlxFontStyles.StrikeOut) != 0) tb.TextDecorations.Add(TextDecorations.Strikethrough);
#endif
            if (Fx.Underline != TFlxUnderline.None) tb.TextDecorations.Add(TextDecorations.Underline);
            if (Fx.Underline == TFlxUnderline.Double || Fx.Underline == TFlxUnderline.DoubleAccounting) tb.TextDecorations.Add(TextDecorations.Baseline);
        }


        /// <summary>
        /// Tries to create a new font given the Excel data.
        /// </summary>
        /// <param name="Fx">FlexCel font with the font information.</param>
        /// <param name="Adj">An adjustment parameter to multiply the FontSize.</param>
        public static Font CreateFont(TFlxFont Fx, real Adj)
        {
            Font Result = new Font();

            Result.Family = new FontFamily(Fx.Name);
            Result.SizeInPix = Fx.Size20 / 20.0 * FlxConsts.PixToPoints * Adj; //fontsize is in pixels, not points
            if ((Fx.Style & TFlxFontStyles.Italic) != 0) Result.Style = FontStyles.Italic; else Result.Style = FontStyles.Normal;
            if ((Fx.Style & TFlxFontStyles.Bold) != 0) Result.Weight = FontWeights.Bold; else Result.Weight = FontWeights.Normal;

            Result.Decorations = new TextDecorationCollection();
            if ((Fx.Style & TFlxFontStyles.StrikeOut) != 0) Result.Decorations.Add(TextDecorations.Strikethrough);
            if (Fx.Underline != TFlxUnderline.None) Result.Decorations.Add(TextDecorations.Underline);
            if (Fx.Underline == TFlxUnderline.Double || Fx.Underline == TFlxUnderline.DoubleAccounting) Result.Decorations.Add(TextDecorations.Baseline);

            Result.Freeze();
            return Result;
        }
    }


    internal class TFontCache : IDisposable
    {
        internal Font GetFont(TFlxFont Fx, real Adj)
        {
            // Not really need to keep a cache in WPF.
            Font NewFont = ExcelFont.CreateFont(Fx, Adj);
            return NewFont;
        }


        
    #region IDisposable Members

        public void Dispose()
        {
            //Not really needed in WPF//;
        }

        #endregion
    }

#else
    /// <summary>
	/// Utility methods to create normal fonts from Excel ones.
	/// </summary>
	public sealed class ExcelFont
	{
		private ExcelFont(){}

        private static Font NewFont(string FontName, float FontSize, FontStyle Style)
        {
#if (MONOTOUCH)
			Font Result = Font.FromName(FontName, FontSize);
#else
            Font Result = new Font(FontName, FontSize, Style);
#endif
            if (FlexCelTrace.HasListeners && !String.Equals(Result.Name.Trim(), FontName.Trim(), StringComparison.InvariantCultureIgnoreCase))
                FlexCelTrace.Write(new TPdfFontNotFoundError(FlxMessages.GetString(FlxErr.ErrFontNotFound, FontName, Result.Name), FontName, Result.Name));
            return Result;
        }

		/// <summary>
		/// Tries to create a new font given the Excel data.
		/// </summary>
		/// <param name="FontName">Name of the font we want to create.</param>
		/// <param name="FontSize">Size of the font.</param>
		/// <param name="Style">Style of the font.</param>
		/// <returns>A new font with the desired parameters.</returns>
        public static Font CreateFont(string FontName, float FontSize, FontStyle Style)
        {
            try
            {
                return NewFont(FontName, FontSize, Style);
            }
            catch (ArgumentException)
            {
            }

            // some fonts like monotype corsiva do not allow to be created with TFontStyle.Regular.
#if(!COMPACTFRAMEWORK && !MONOTOUCH)

            FontFamily[] Families = FontFamily.Families;
            try
            {
                for (int i = Families.Length - 1; i >= 0; i--)
                {
                    FontFamily ff = Families[i];
                    if (String.Equals(ff.Name, FontName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (!ff.IsStyleAvailable(Style))
                            Style |= FontStyle.Italic;
                        if (!ff.IsStyleAvailable(Style))
                            Style &= ~FontStyle.Italic;
                        if (!ff.IsStyleAvailable(Style))
                            Style |= FontStyle.Bold;
                        if (!ff.IsStyleAvailable(Style))
                            Style &= ~FontStyle.Bold;
                        if (!ff.IsStyleAvailable(Style))
                            Style &= FontStyle.Regular;  //No more tries...

                        try
                        {
                            return NewFont(FontName, FontSize, Style);
                        }
                        catch (Exception ex)
                        {
                            FlxMessages.ThrowException(ex, FlxErr.ErrFontNotSupported, FontName, FontSize, Style, ex.Message);
                        }
                    }
                }
            }
            finally
            {
                foreach (FontFamily ff in Families) ff.Dispose();
            }
#endif

            //Font is not available.
            return NewFont("Arial", FontSize, FontStyle.Regular);

        }

		/// <summary>
		/// Converts between a standard FontStyle and a TFlxFont.
		/// </summary>
		/// <param name="Fx">FlexCel font to convert.</param>
		/// <returns>A similar FontStyle.</returns>
		public static FontStyle ConvertFontStyle(TFlxFont Fx)
		{
			FontStyle DrawFontStyle = FontStyle.Regular;
			if (Fx.Underline != TFlxUnderline.None)
				DrawFontStyle |= FontStyle.Underline;

			if ((Fx.Style & TFlxFontStyles.Bold) != 0) DrawFontStyle |= FontStyle.Bold;
			if ((Fx.Style & TFlxFontStyles.Italic) != 0) DrawFontStyle |= FontStyle.Italic;
			if ((Fx.Style & TFlxFontStyles.StrikeOut) != 0) DrawFontStyle |= FontStyle.Strikeout;

			return DrawFontStyle;
		}
	}


	internal class TFontInfo: IDisposable
	{
		internal Font FFont;
		internal int TimesUsed;

		internal TFontInfo(Font aFont)
		{
			FFont=aFont;
		}

		#region IDisposable Members
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (FFont!=null) FFont.Dispose();
				FFont = null;
			}
		}

		#endregion

	}

    internal struct TFontDesc
    {
        internal string Name;
        internal real Size;
        internal FontStyle Style;

        internal TFontDesc(real aSize, string aName, FontStyle aStyle)
        {
            Size = aSize;
            Name = aName;
            Style = aStyle;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is TFontDesc)) return false;
            TFontDesc o2 = (TFontDesc)obj;
            return Size == o2.Size && Name == o2.Name && Style == o2.Style;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHash(Size.GetHashCode(), Name.GetHashCode(), ((int)Style).GetHashCode());
        }
    }

    internal sealed class TFontList : Dictionary<TFontDesc, TFontInfo>
    {
        public TFontList(): base()
        {
        }
    }


	internal class TFontCache: IDisposable
	{
		TFontList Fonts;
		private const int MaxCacheSize = 20;

		internal TFontCache()
		{
			Fonts = new TFontList();
		}

		internal void ClearCache()
		{
			foreach(TFontInfo f in Fonts.Values) f.Dispose();
			Fonts.Clear();
		}

		internal Font GetFont(TFlxFont aFont, real SizeAdj)
		{
			
			real FontSize = (real)(aFont.Size20 / 20f * SizeAdj);
			if (FontSize < 1) FontSize = 1;

			FontStyle Style = ExcelFont.ConvertFontStyle(aFont);

            TFontDesc key = new TFontDesc(FontSize, aFont.Name, Style);
			TFontInfo f = null;
			if (Fonts.TryGetValue(key, out f))
			{
				f.TimesUsed++;
				return f.FFont;
			}

			//Don't do this!. We need to dispose only fonts that are not still being used.
			//if (Fonts.Count>MaxCacheSize)
			//	ClearCache();
			
			Font NewFont = ExcelFont.CreateFont(aFont.Name, FontSize, Style);
			Fonts[key] = new TFontInfo(NewFont);
			return NewFont;
		}


		#region IDisposable Members
		public void Dispose()
		{
			ClearCache();
            GC.SuppressFinalize(this);
        }
		

		#endregion
	}
#endif
}
