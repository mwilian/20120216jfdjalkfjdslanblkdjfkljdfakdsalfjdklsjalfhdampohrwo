using System;

using FlexCel.Core;

#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	using SizeF = System.Windows.Size;
	using real = System.Double;
	using System.Windows.Media;
	#else
	using real = System.Single;
	using System.Drawing;
	#endif
#endif
namespace FlexCel.Render
{
	/// <summary>
	/// Class for calculating widths and heights of strings. UCS32/surrogate enabled.
	/// </summary>
	internal sealed class RenderMetrics
	{
		private RenderMetrics(){}

        internal static void FitOneLine(IFlxGraphics Canvas, TFontCache FontCache, real Zoom100, TRichString Text, Font AFont, real w, int idx, ref int fit, ref real StrWidth)
        {
            if (StrWidth > w)
            {
                real md = 0;
              
                fit = Math.Max((int)Math.Round(fit * w / StrWidth), 1);
				CharUtils.SameOrMore(Text.Value, idx, ref fit);
                StrWidth = CalcTextExtent(Canvas, FontCache, Zoom100, AFont, Text.Substring(idx, fit), out md).Width;

                if (StrWidth > w)
                {
                    while (fit > 1 && StrWidth > w) 
					{ 
						fit--; 
						CharUtils.SameOrLess(Text.Value, idx, ref fit);
						StrWidth = CalcTextExtent(Canvas, FontCache, Zoom100, AFont, Text.Substring(idx, fit), out md).Width; 
					}
                }
                else
                {
                    while (StrWidth < w) //Will not overflow, as wc>w for fit+idx=MaxLength.
                    {
						int fit1 = fit + 1;
						CharUtils.SameOrMore(Text.Value, idx, ref fit1);
						StrWidth = CalcTextExtent(Canvas, FontCache, Zoom100, AFont, Text.Substring(idx, fit1), out md).Width;
                        if (StrWidth < w) fit = fit1;
                    }
                }
            }
        }

        internal static void SplitText(IFlxGraphics Canvas, TFontCache FontCache, real Zoom100,
			TRichString Text, Font AFont, real w, TXRichStringList TextLines, out SizeF TextExtent, bool Vertical, TFloatList MaxDescent, TAdaptativeFormats AdaptFormat)
		{
			TextLines.Clear();
			MaxDescent.Clear();

			int idx = 0;
			int fit = 0;


			TextExtent = new SizeF(0, 0);
			if (w <= 0 || Text.Value.Length <= 0) return;

			while (idx < Text.Value.Length)
			{
				int Enter = Text.Value.IndexOf((char)0x0A, idx);
				int MaxLength = Text.Value.Length - idx;
				if (Enter >= 0) MaxLength = Enter - idx;

				if (Vertical) 
				{
					fit = 1;
					CharUtils.SameOrMore(Text.Value, idx, ref fit);
				}
				else
				{
					//Not a really efficient way, but...
					//First Guess. whole string.
					fit = MaxLength;

					real md;
					real wc = CalcTextExtent(Canvas, FontCache, Zoom100, AFont, Text.Substring(idx, fit), out md).Width;
					FitOneLine(Canvas, FontCache, Zoom100, Text, AFont, w, idx, ref fit, ref wc);

					if (fit <= 0) fit = 1; //avoid infinite loop
					CharUtils.SameOrMore(Text.Value, idx, ref fit);


					if (Text.Value.IndexOf(' ', idx, fit) >= 0)
					{
						int minfit = 1;
						CharUtils.SameOrMore(Text.Value, idx, ref minfit);
						while (fit > minfit && fit < MaxLength && Text.Value[idx + fit] != ' ') fit--;
					}
					while (fit < MaxLength && Text.Value[idx + fit] == ' ') fit++;
					//No need to adjust fit for surrogates. it will always be at the start of one.
				}

				//int Enter=Text.Value.IndexOf((char)0x0A, idx, fit);
				//if (Enter>0) fit=Enter-idx;
                int TextLen = Math.Min(MaxLength, fit);
				TextLines.Add(new TXRichString(Text.Substring(idx, TextLen),true, 0, 0, TAdaptativeFormats.CopyTo(AdaptFormat, idx, TextLen)));

				if (fit + idx < Text.Value.Length && Text.Value[idx + fit] == (char)0x0A)
				{
					TextLines[TextLines.Count-1].Split = false;
					if (idx+fit < Text.Value.Length -1) //An Enter at the end behaves different, it means we have a new empty line.
						idx++;
				}
				if (fit + idx >= Text.Value.Length)
				{
					TextLines[TextLines.Count-1].Split = false;					
				}

				idx += fit;

				//Recalculate dx
				real mdx = 0;
				TRichString sx = TextLines[TextLines.Count - 1].s;
				SizeF bSize = CalcTextExtent(Canvas, FontCache, Zoom100, AFont, sx, out mdx);
				if (bSize.Height == 0)
				{
					real mdx3;
					SizeF bSize3 = CalcTextExtent(Canvas, FontCache, Zoom100, AFont, new TRichString("M"), out mdx3);
					bSize.Height = bSize3.Height;
				}

				SizeF bSize2 = bSize;
				if (sx != null && sx.Length>0 && sx.ToString()[sx.Length-1] == ' ') //This is to right align line with spaces at the end.
				{
					real mdx2;
					bSize2 = CalcTextExtent(Canvas, FontCache, Zoom100, AFont, sx.RightTrim(), out mdx2);
				}
				TextLines[TextLines.Count - 1].XExtent = bSize2.Width;
				TextLines[TextLines.Count - 1].YExtent = bSize.Height; // not bSize2.Height; This might be even 0.
				
				MaxDescent.Add(mdx);
				if (TextExtent.Width < bSize.Width) TextExtent.Width = bSize.Width;
				TextExtent.Height += bSize.Height;
			} //while
		}

		internal static SizeF CalcTextExtent(IFlxGraphics Canvas, TFontCache FontCache, real Zoom100,
			Font AFont, TRichString Text, out real MaxDescent)
		{
			MaxDescent = Canvas.FontDescent(AFont);
			if (Text.RTFRunCount == 0)
			{
				return Canvas.MeasureStringEmptyHasHeight(Text.Value, AFont);
			}

			real x = 0;
			real y = 0;
			SizeF Result;

			Result = Canvas.MeasureString(Text.Value.Substring(0, Text.RTFRun(0).FirstChar), AFont, new TPointF(0, 0));
			x = Result.Width; y = Result.Height;


			for (int i = 0; i < Text.RTFRunCount - 1; i++)
			{
				TFlxFont Fx = Text.GetFont(Text.RTFRun(i).FontIndex);
				TSubscriptData Sub = new TSubscriptData(Fx.Style);                            
				Font MyFont = FontCache.GetFont(Fx, Zoom100 * Sub.Factor);
			{
				int Start = Text.RTFRun(i).FirstChar;
				if (Start >= Text.Length) Start = Text.Length;
				int Len = Text.RTFRun(i + 1).FirstChar;
				if (Len >= Text.Length) Len = Text.Length;
				Len -= Start;
                if (Len < 0) continue;  //wrong file, (i+1)FirstChar < (i).FirstChar
				Result = Canvas.MeasureString(Text.Value.Substring(Start, Len), MyFont, new TPointF(0, 0));
				x += Result.Width; y = Math.Max(y + Sub.Offset(Canvas, MyFont), Result.Height);
				MaxDescent = Math.Max(MaxDescent, Canvas.FontDescent(MyFont) + Sub.Offset(Canvas, MyFont));
			}
			}


			TFlxFont Fy = Text.GetFont(Text.RTFRun(Text.RTFRunCount - 1).FontIndex);
			TSubscriptData Suby = new TSubscriptData(Fy.Style);                            
			Font MyFont2 = FontCache.GetFont(Fy, Zoom100 * Suby.Factor);
		{
			int Start = Text.RTFRun(Text.RTFRunCount - 1).FirstChar;
			if (Start >= Text.Length) Start = Text.Length;
			Result = Canvas.MeasureStringEmptyHasHeight(Text.Value.Substring(Start), MyFont2);
			x += Result.Width; y = Math.Max(y + Suby.Offset(Canvas, MyFont2), Result.Height);
			MaxDescent = Math.Max(MaxDescent, Canvas.FontDescent(MyFont2) + Suby.Offset(Canvas, MyFont2));
		}

			return new SizeF(x, y);
		}


	}
}
