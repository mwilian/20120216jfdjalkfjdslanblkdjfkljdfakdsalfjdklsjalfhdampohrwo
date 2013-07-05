using System;
using System.Globalization;

using FlexCel.Core;
using System.Collections.Generic;

#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	using SizeF = System.Windows.Size;
	using System.Windows.Media;
	using real = System.Double;
	#else
	using System.Drawing;
	using real = System.Single;
	#endif
#endif

namespace FlexCel.Render
{
	#region TextWriter
	internal sealed class TextPainter
	{
        private TextPainter() { }
		#region Utiilties
		private static bool Intersect(RectangleF Rect1, RectangleF Rect2, out RectangleF OutRect)
		{
			OutRect = RectangleF.FromLTRB(
				Math.Max(Rect1.Left, Rect2.Left),
				Math.Max(Rect1.Top, Rect2.Top),
				Math.Min(Rect1.Right, Rect2.Right),
				Math.Min(Rect1.Bottom, Rect2.Bottom)
				);
			return OutRect.Left < OutRect.Right-0.001 && OutRect.Top < OutRect.Bottom-0.001;
		}

		private static Color GetColor(ExcelFile Workbook, TExcelColor aColor)
		{
            return aColor.ToColor(Workbook, Color.Black);
		}

		#endregion

		#region BIDI
		internal static bool IsLeftToRight(string Text)
		{
			for (int i=0; i< Text.Length; i++)
			{
				sbyte l = BidiReference.Direction(Text[i]);
				if (l<= BidiReference.LRO) return false;
				if (l<= BidiReference.RLO) return true;
			}
			return false;
		}

		internal static TRichString ArabicShape(TRichString Text, bool ReverseRightToLeftStrings)
		{
			if (!ReverseRightToLeftStrings || Text == null || Text.Length<2) return Text;

			ArabicShaping Ashaping = new ArabicShaping(ArabicShaping.LENGTH_GROW_SHRINK+
				ArabicShaping.TEXT_DIRECTION_LOGICAL+ ArabicShaping.LETTERS_SHAPE);
			string SResult = Ashaping.shape(Text.Value); 

			if (String.Equals(Text.Value, SResult, StringComparison.InvariantCulture))
				return Text;
			return new TRichString(SResult);  
		}

		internal static TRichString GetVisualString(TRichString Text, bool ReverseRightToLeftStrings)
		{
			if (!ReverseRightToLeftStrings || Text == null || Text.Length<2) return Text;
			BidiReference bref = new BidiReference(Text.Value);
			int[] lines = {Text.Length};
			int[] ReorderArray = bref.getReordering(lines);

			bool ToLeft = true;
			for (int i=0; i< ReorderArray.Length; i++)
			{
				if (ReorderArray[i]!=i)
				{
					ToLeft = false;
					break;
				}
			}

			if (ToLeft) return Text;
			char[] Result = new char[ReorderArray.Length];
			for (int i=0; i< Result.Length; i++)
				Result[i] = Text[ReorderArray[i]];

			return new TRichString(new string(Result));  

		}
		#endregion

		#region TextCoords
		internal static void RelocateBox(ref RectangleF ContainingRect, real[]X, real[]Y, real ofsx, real ofsy)
		{
			ContainingRect.Offset(ofsx, ofsy);
			for (int i = 0; i < X.Length; i++)
			{
				X[i] += ofsx;
				Y[i] += ofsy;
			}
		}


        private static real CalcMaxTextWidth(ref RectangleF CellRect, real Clp, real Alpha, real TextHeight)
        {
            if (Alpha == 0) return CellRect.Right - CellRect.Left - 2 * Clp;
            else
            {
                real SinAlpha = (real)Math.Sin(Alpha * Math.PI / 180); real CosAlpha = (real)Math.Cos(Alpha * Math.PI / 180);
                return (CellRect.Height - 2 * Clp - TextHeight * CosAlpha) / Math.Abs(SinAlpha);
                /*No. Rotated text only considers row height, and never cell width.
                if (Math.Abs(CosAlpha)> 1e-5)
                {
                    real Wr2 = (CellRect.Width - 2 * Clp - TextExtent.Height * SinAlpha);
                    Wr = Math.Min(Wr, Wr2 /Math.Abs(CosAlpha));
                }*/
            }
        }

		internal static void CalcTextBox(IFlxGraphics Canvas, TFontCache FontCache, real Zoom100, RectangleF CellRect, real Clp, bool MultiLine, real Alpha, bool Vertical, TRichString OutText, Font AFont, TAdaptativeFormats AdaptativeFormats, out SizeF TextExtent, out TXRichStringList TextLines, out TFloatList MaxDescent)
		{
			TextExtent = new SizeF(0, 0);
			TextLines = new TXRichStringList();

			MaxDescent = new TFloatList();
            real MaxTextWidth;

			if (MultiLine || Vertical)
			{
                real Md;
                real TextHeight = RenderMetrics.CalcTextExtent(Canvas, FontCache, Zoom100, AFont, new TRichString("M"), out Md).Height; //not perfect since the rich string might have different fonts, but a good approx.
                MaxTextWidth = CalcMaxTextWidth(ref CellRect, Clp, Alpha, TextHeight);

				RenderMetrics.SplitText(Canvas, FontCache, Zoom100, OutText, AFont, MaxTextWidth, TextLines, out TextExtent, Vertical, MaxDescent, AdaptativeFormats);
			}
			else
			{
				TextLines.Add(new TXRichString(OutText, false, 0, 0, AdaptativeFormats));
				real mdx = 0;
				TextExtent = RenderMetrics.CalcTextExtent(Canvas, FontCache, Zoom100, AFont, OutText, out mdx);
				TextLines[0].XExtent = TextExtent.Width;
				TextLines[0].YExtent = TextExtent.Height;
                MaxTextWidth = CalcMaxTextWidth(ref CellRect, Clp, Alpha, TextExtent.Height);
                MaxDescent.Add(mdx);
			}

            if (AdaptativeFormats != null && AdaptativeFormats.WildcardPos >= 0)
            {
                MaxTextWidth -= 2 * Clp; //Add some extra clipping so the text doesn't go through the line.
                if (Vertical)
                {
                }
                else
                {
                    AddWildcard(Canvas, FontCache, Zoom100, AFont, MaxTextWidth, TextLines, ref TextExtent);
                }
            }

		}

        private static void AddWildcard(IFlxGraphics Canvas, TFontCache FontCache, real Zoom100,
            Font AFont, real MaxTextWidth, TXRichStringList TextLines, ref SizeF TextExtent)
        {
            TXRichString LineWithWildcard = null;
            foreach (TXRichString line in TextLines)
            {
                if (line.s == null || line.s.RTFRunCount > 0) return; //wildcards only apply to non formatted lines.
                if (line.AdaptFormat != null && line.AdaptFormat.WildcardPos >= 0)
                {
                    LineWithWildcard = line;
                    break;
                }
            }

            if (LineWithWildcard == null) return;
            AddWilcardtoLine(Canvas, FontCache, Zoom100, AFont, MaxTextWidth, ref TextExtent, LineWithWildcard);
        }

        internal static void AddWilcardtoLine(IFlxGraphics Canvas, TFontCache FontCache, real Zoom100, Font AFont, real MaxTextWidth, ref SizeF TextExtent, TXRichString LineWithWildcard)
        {
            if (LineWithWildcard == null || LineWithWildcard.AdaptFormat == null) return;
            int WildcardPos = LineWithWildcard.AdaptFormat.WildcardPos;
            if (WildcardPos < 0) return;

            string s1 = LineWithWildcard.s.ToString();
            string WildcardChar = String.Empty + s1[WildcardPos];
            if (CharUtils.IsSurrogatePair(s1, WildcardPos) && WildcardPos + 1 < s1.Length) WildcardChar += s1[WildcardPos + 1];
            string sOrg = s1;
            s1 = s1.Remove(WildcardPos, WildcardChar.Length);//consider the case the string has 0 wildcards.
            string sOld = s1;
            real wc;
            real OldWc = LineWithWildcard.XExtent;
            do
            {
                real md;
                wc = RenderMetrics.CalcTextExtent(Canvas, FontCache, Zoom100, AFont, new TRichString(s1), out md).Width;
                if (wc < MaxTextWidth)
                {
                    sOld = s1;
                    OldWc = wc;
                    s1 = s1.Insert(WildcardPos, WildcardChar);
                }
            }
            while (wc < MaxTextWidth);

            LineWithWildcard.s = new TRichString(sOld);
            if (sOrg.Length > sOld.Length) LineWithWildcard.AdaptFormat.RemovedPosition(WildcardPos, sOrg.Length - sOld.Length); //a line with a wildcard can have 0 characters in the wildcard.
            if (sOrg.Length < sOld.Length) LineWithWildcard.AdaptFormat.InsertedPosition(WildcardPos, sOld.Length - sOrg.Length);
            LineWithWildcard.XExtent = OldWc;
            if (TextExtent.Width < OldWc) TextExtent.Width = OldWc;
        }
 

		internal static RectangleF CalcTextCoords(out real[] X, out real[] Y, TRichString OutText, TVAlign VAlign, 
			ref THAlign HAlign, real Indent, real Alpha, RectangleF CellRect, real Clp, SizeF TextExtent,
			bool FmtGeneral, bool Vertical, real SinAlpha, real CosAlpha, TXRichStringList TextLines, double Linespacing0, TVFlxAlignment VJustify)
		{
            double Linespacing = VJustify == TVFlxAlignment.distributed || VJustify == TVFlxAlignment.justify ? 1.0 : Linespacing0;
            
            bool TextLeftToRight = IsLeftToRight(OutText.ToString());

			X = new real[TextLines.Count]; Y =  new real[TextLines.Count]; 
			
			if (Alpha == 0)  //optimize the most common case
			{
				return CalcTextCoordsAlpha0(X, Y, VAlign, ref HAlign, Indent, ref CellRect, Clp, ref TextExtent, FmtGeneral, Vertical, TextLeftToRight, TextLines, Linespacing);
			}

			else
				// There is rotation. Now values of TextExtent are not ok, since lines will stack.for example ///////  Each line has the same y, and the x is calculated as height/cosAlpha
				return CalcTextCoordsRotated(X, Y, VAlign, HAlign, Alpha, ref CellRect, Clp, FmtGeneral, SinAlpha, CosAlpha, TextLines, Linespacing);
		}

		private static RectangleF CalcTextCoordsAlpha0(real[] X, real[] Y, TVAlign VAlign, ref THAlign HAlign, real Indent, ref RectangleF CellRect, real Clp, ref SizeF TextExtent, bool FmtGeneral, bool Vertical, bool IsLeftToRight, TXRichStringList TextLines, double Linespacing)
		{
			RectangleF ContainingRect;
			real AcumY = 0;
			real MinX = 0;
            real TextHeight = Y.Length < 1 ? TextExtent.Height : (float)(TextLines[Y.Length - 1].YExtent + (TextExtent.Height - TextLines[Y.Length - 1].YExtent) * Linespacing); //Last line doesn't add linespace.
			for (int i = 0; i < X.Length; i++)
			{
				switch (VAlign)
				{
					case TVAlign.Top: Y[i] = AcumY + CellRect.Top + Clp + TextLines[i].YExtent; break; //Used to be "CellRect.Top + Clp + TextExtent.Height * Lines" instead, but it goes too low.
					case TVAlign.Center: Y[i] = AcumY + (CellRect.Top + CellRect.Bottom - TextHeight) / 2 + TextLines[i].YExtent; break;
					default: Y[i] = AcumY + CellRect.Bottom - Clp - TextHeight + TextLines[i].YExtent; break;
				} //case

				AcumY += (float)(TextLines[i].YExtent * Linespacing);


				THAlign RHAlign = HAlign;
				if (FmtGeneral)
				{
					if (IsLeftToRight) { RHAlign = THAlign.Right; HAlign = RHAlign; }
					if (Vertical) { RHAlign = THAlign.Center; HAlign = RHAlign; }
				}
				switch (RHAlign)
				{
					case THAlign.Right: X[i] = CellRect.Right - Clp - TextLines[i].XExtent - Indent; break;
					case THAlign.Center: X[i] = (CellRect.Left + CellRect.Right - TextLines[i].XExtent) / 2; break;
					default: X[i] = CellRect.Left + Clp + 1f * FlexCelRender.DispMul / 100f + Indent; break;
				} //case

				if (i == 0 || MinX > X[i]) MinX = X[i];
			}
			ContainingRect = new RectangleF(MinX, Y[0] - TextLines[0].YExtent, TextExtent.Width, TextHeight);
			return ContainingRect;
		}

        private static RectangleF CalcTextCoordsRotated(real[] X, real[] Y, TVAlign VAlign, THAlign HAlign, real Alpha, ref RectangleF CellRect, real Clp, bool FmtGeneral, real SinAlpha, real CosAlpha, TXRichStringList TextLines, double Linespacing)
		{
			//General horiz align depends on the rotation
			THAlign RHAlign = HAlign;
			if (FmtGeneral)
				if (Alpha == 90) RHAlign = THAlign.Right;
				else if (Alpha > 0) RHAlign = THAlign.Left;
				else if (Alpha != -90) RHAlign = THAlign.Right;
				else RHAlign = THAlign.Left;


			/*
			 * This is quite complex because each box of text has 4 corners. We always write on left-bottom coordinate, but to keep
			 * calculations easy, we will calculate the left-bottom for alpha > 0, and left-top for alpha < 0
			 *            ___                   ___x
			 * (Alpha>0) /  /      (Alpha < 0)  \  \ 
			 *          /  /      	             \  \
			 *          --x       			      ---
			 * 
			 * We will then adjust the x coord for alpha < 0 to be left-bottom.
			 */
            
			real xdiff = 0;
			real XOfs = 0;
			real GlobalMinBoxX =0, GlobalMaxBoxX = 0;
			
			for (int i = 0; i < TextLines.Count; i++)
			{
				real hcosAlpha = TextLines[i].YExtent * CosAlpha;
				real wsinAlpha = TextLines[i].XExtent * SinAlpha;
				real dy = Math.Abs(wsinAlpha) + hcosAlpha;

				switch (VAlign)
				{
					case TVAlign.Top:
						Y[i] = CellRect.Top + Clp + dy;
						break;
					case TVAlign.Center:
						Y[i] = (CellRect.Top + CellRect.Bottom + dy) / 2;
						break;
					default:
						Y[i] = CellRect.Bottom - Clp;
						break;
				} //case

				if (Alpha < 0) Y[i] -= dy; 


				//Now accumulate in x
				real wcosAlpha = TextLines[i].XExtent * CosAlpha;
				real hsinAlpha = TextLines[i].YExtent * SinAlpha;

				if (i > 0)
				{
					int z = Alpha>0? i: i-1;
					real xinc = TextLines[z].YExtent / SinAlpha;  //xinc here is what we need to stack boxes one after the other, keeping the same Y.
					XOfs += xinc;
				}
				
				X[i] = XOfs;
				if (i > 0) X[i] += (Y[0] - Y[i])  * CosAlpha / SinAlpha; //if both Y are not aligned, we need to substract that offset.

				switch (RHAlign)
				{
					case THAlign.Right:
						real MaxBoxX = X[i] + wcosAlpha;
						if (MaxBoxX > GlobalMaxBoxX) 
						{
							GlobalMaxBoxX = MaxBoxX;
							xdiff = CellRect.Right - Clp - MaxBoxX;
						}

						break;
					case THAlign.Center:
						real cMaxBoxX = X[i] + wcosAlpha;
						if (cMaxBoxX > GlobalMaxBoxX) 
						{
							GlobalMaxBoxX = cMaxBoxX;
						}
						real cMinBoxX = X[i] - Math.Abs(hsinAlpha);
						if (cMinBoxX < GlobalMinBoxX) 
						{
							GlobalMinBoxX = cMinBoxX;
						}
						if ( i == TextLines.Count - 1)
						{
							if (Alpha > 0)
							{
								real h0sinAlpha = TextLines[0].YExtent * SinAlpha;
								xdiff = (CellRect.Left + CellRect.Right - (GlobalMaxBoxX - GlobalMinBoxX)) / 2 + h0sinAlpha;
							}
							else
							{
								real w0cosAlpha = TextLines[0].XExtent * CosAlpha;
								xdiff = (CellRect.Left + CellRect.Right + (GlobalMaxBoxX - GlobalMinBoxX)) / 2 - w0cosAlpha;
							}
						}

						break;
					default:
						real MinBoxX = X[i] - Math.Abs(hsinAlpha);
						if (MinBoxX < GlobalMinBoxX) 
						{
							GlobalMinBoxX = MinBoxX;
							xdiff = CellRect.Left + Clp - MinBoxX;
						}
						break;
				} //case

			}

			real MinX = 0, MinY = 0, MaxX = 0, MaxY = 0;
			//second pass. adjust offsets. if alpha < 0, adjust x to be left-bottom, to write the text.
			for (int i = 0; i < TextLines.Count; i++)
			{
				X[i] += xdiff;

				real wcosAlpha = TextLines[i].XExtent * CosAlpha;
				real hsinAlpha = TextLines[i].YExtent * SinAlpha;
				real hcosAlpha = TextLines[i].YExtent * CosAlpha;
				real wsinAlpha = TextLines[i].XExtent * SinAlpha;

				if (i == 0)
				{
					CalcBox(X[i], Y[i], Alpha, wcosAlpha, hsinAlpha, hcosAlpha, wsinAlpha, out MinX, out MinY, out MaxX, out MaxY);
				}
				else
				{
					real aMinX, aMinY, aMaxX, aMaxY; 
					CalcBox(X[i], Y[i], Alpha, wcosAlpha, hsinAlpha, hcosAlpha, wsinAlpha, out aMinX, out aMinY, out aMaxX, out aMaxY);
					MinX = Math.Min(MinX, aMinX);
					MaxX = Math.Max(MaxX, aMaxX);
					MinY = Math.Min(MinY, aMinY);
					MaxY = Math.Max(MaxY, aMaxY);
				}

				if (Alpha < 0)
				{
					Y[i] += hcosAlpha;
					X[i] += hsinAlpha; //adjust because we are not using the same x to write the text.
				}


			}
			return RectangleF.FromLTRB(MinX, MinY, MaxX, MaxY);
		}

		private static void CalcBox(real X, real Y, real Alpha, real wcosAlpha, real hsinAlpha, real hcosAlpha, real wsinAlpha, out real MinX, out real MinY, out real MaxX, out real MaxY)
		{
			MinX = X - Math.Abs(hsinAlpha);
			MaxX = X + wcosAlpha;

			if (Alpha > 0)
			{
				MinY = Y - wsinAlpha - hcosAlpha;
				MaxY = Y;
			}
			else
			{
				MinY = Y;
				MaxY = Y - wsinAlpha + hcosAlpha;
			}
		}
		#endregion

		#region DrawText
        private static void WriteJustText(ExcelFile Workbook, IFlxGraphics Canvas, TFontCache FontCache, real Zoom100, bool ReverseRightToLeftStrings, Font AFont, Color AFontColor, real y, real SubOfs, RectangleF CellRect, real Clp, TRichString OutText, real XExtent, real MaxDescent, bool Distributed, TAdaptativeFormats AdaptativeFormats)
        {
            List<TRichString> Words = new List<TRichString>();
            OutText = GetVisualString(OutText, ReverseRightToLeftStrings);
            string s = OutText.Value;
            int p = 0; int p1 = 0;
            while ((p1 = s.IndexOf(' ', p)) >= 0)
            {
                Words.Add(OutText.Substring(p, p1 - p + 1));
                p = p1 + 1;
            }

            if (p < s.Length) Words.Add(OutText.Substring(p));

            real md;
            real wc = XExtent; // CalcTextExtent(AFont, OutText, out md).Width;

            real Spaces = 0;
            if (Words.Count - 1 > 0) Spaces = (CellRect.Width - 2 - 2 * Clp - wc) / (Words.Count - 1);

            real x = CellRect.Left + 1 + Clp;
            if (Words.Count == 1 && Distributed)  //Center when it is one word
            {
                x = (CellRect.Left + CellRect.Right - 2 - 2 * Clp - wc) / 2;
            }
            for (int i = 0; i < Words.Count; i++)
            {
                TRichString wo = Words[i];
                WriteText(Workbook, Canvas, FontCache, Zoom100, AFont, AFontColor, x, y, SubOfs, wo, 0, MaxDescent, AdaptativeFormats);
                x += RenderMetrics.CalcTextExtent(Canvas, FontCache, Zoom100, AFont, wo, out md).Width;
                if (Spaces > 0) x += Spaces;
            }

        }

        internal static void WriteText(ExcelFile Workbook, IFlxGraphics Canvas, TFontCache FontCache, real Zoom100, Font AFont, Color AFontColor, real x, real y, real SubOfs, TRichString OutText, real Alpha, real MaxDescent, TAdaptativeFormats AdaptativeFormats)
		{
			if (OutText.Length == 0) return;

			if (Alpha != 0) Canvas.SaveTransform();
			try
			{
				if (Alpha != 0)
				{
					Canvas.Rotate(x, y, Alpha);
				}

				using (Brush TextBrush = new SolidBrush(AFontColor))
				{
                    if (OutText.RTFRunCount == 0)
                    {
                        if (AdaptativeFormats == null || AdaptativeFormats.Separators == null || AdaptativeFormats.Separators.Length == 0) //formats are only applied to non rich text cells.
                        {
                            Canvas.DrawString(OutText.Value, AFont, TextBrush, x, y + SubOfs);
                        }
                        else
                        {
                            DrawAdaptativeString(Canvas, AdaptativeFormats, OutText.Value, AFont, TextBrush, x, y + SubOfs);
                        }
                    }
                    else
                    {
                        SizeF Result;

                        string s1 = OutText.Value.Substring(0, OutText.RTFRun(0).FirstChar);
                        if (s1.Length > 0)
                        {
                            Result = Canvas.MeasureString(s1, AFont, new TPointF(x, y));
                            Canvas.DrawString(s1, AFont, TextBrush, x, y + SubOfs - (MaxDescent - Canvas.FontDescent(AFont)));
                            x += Result.Width;
                        }

                        for (int i = 0; i < OutText.RTFRunCount - 1; i++)
                        {
                            TFlxFont Fx = OutText.GetFont(OutText.RTFRun(i).FontIndex);
                            TSubscriptData Sub = new TSubscriptData(Fx.Style);
                            Font MyFont = FontCache.GetFont(Fx, Zoom100 * Sub.Factor);
                            {
                                using (Brush MyBrush = new SolidBrush(GetColor(Workbook, Fx.Color)))
                                {
                                    int Start = OutText.RTFRun(i).FirstChar;
                                    if (Start >= OutText.Length) Start = OutText.Length;
                                    int Len = OutText.RTFRun(i + 1).FirstChar;
                                    if (Len >= OutText.Length) Len = OutText.Length;
                                    Len -= Start;

                                    string s2 = OutText.Value.Substring(Start, Len);
                                    Result = Canvas.MeasureString(s2, MyFont, new TPointF(x, y));
                                    Canvas.DrawString(s2, MyFont, MyBrush, x, y + Sub.Offset(Canvas, MyFont) - (MaxDescent - Canvas.FontDescent(MyFont)));
                                    x += Result.Width;
                                }
                            }
                        }
                        TFlxFont Fy = OutText.GetFont(OutText.RTFRun(OutText.RTFRunCount - 1).FontIndex);
                        TSubscriptData Suby = new TSubscriptData(Fy.Style);
                        Font MyFont2 = FontCache.GetFont(Fy, Zoom100 * Suby.Factor);
                        {
                            using (Brush MyBrush = new SolidBrush(GetColor(Workbook, Fy.Color)))
                            {
                                int Start = OutText.RTFRun(OutText.RTFRunCount - 1).FirstChar;
                                if (Start >= OutText.Length) Start = OutText.Length;

                                string s3 = OutText.Value.Substring(Start);
                                Result = Canvas.MeasureString(s3, MyFont2, new TPointF(x, y));
                                Canvas.DrawString(s3, MyFont2, MyBrush, x, y + Suby.Offset(Canvas, MyFont2) - (MaxDescent - Canvas.FontDescent(MyFont2)));
                            }
                        }
                    }
				}
			}
			finally
			{
				if (Alpha != 0) Canvas.ResetTransform();
			}
		}

        private static void DrawAdaptativeString(IFlxGraphics Canvas, TAdaptativeFormats AdaptativeFormats, string Text, Font AFont, Brush TextBrush, real x, real y)
        {
            int start = 0;
            foreach (TCharAndPos cp in AdaptativeFormats.Separators)
            {
                if (cp.Pos >= Text.Length) break; //just a security measure.
                string SubText = Text.Substring(start, cp.Pos - start);
                Canvas.DrawString(SubText, AFont, TextBrush, x, y);
                start = cp.Pos + 1;
                x += Canvas.MeasureString(SubText + Text[cp.Pos], AFont).Width;
                if (start-1 < Text.Length && CharUtils.IsSurrogatePair(Text, start-1)) start++;
            }

            if (start < Text.Length) Canvas.DrawString(Text.Substring(start), AFont, TextBrush, x, y);
        }

		internal static void DrawRichText(ExcelFile Workbook, IFlxGraphics Canvas, TFontCache FontCache, real Zoom100, bool ReverseRightToLeftStrings, 
			ref RectangleF CellRect, ref RectangleF PaintClipRect, ref RectangleF TextRect, ref RectangleF ContainingRect,
			real Clp, THFlxAlignment HJustify, TVFlxAlignment VJustify, real Alpha, 
			Color DrawFontColor, TSubscriptData SubData, TRichString OutText, SizeF TextExtent, 
			TXRichStringList TextLines, 
			Font AFont, TFloatList MaxDescent, real[] X, real[] Y)
		{

			RectangleF FinalRect = CellRect;
			if (ContainingRect.Right > PaintClipRect.Left &&
				ContainingRect.Left < PaintClipRect.Right &&
				Intersect(TextRect, PaintClipRect, out FinalRect))
			{
				real dy = 0;
				int TextLinesCount = TextLines.Count;

				if (VJustify == TVFlxAlignment.justify || VJustify == TVFlxAlignment.distributed)
				{
					if (TextLinesCount > 1)
						dy = (CellRect.Height - TextExtent.Height - 2 * Clp) / (TextLinesCount - 1);
				}

                float eps = 0.0001F; //Tolerance for the comparison, these are floats, not doubles, and we get rounding errors soon.
                if (TextRect.Left - eps <= ContainingRect.Left && TextRect.Right + eps >= ContainingRect.Right && TextRect.Top - eps <= ContainingRect.Top && TextRect.Bottom + eps >= ContainingRect.Bottom)
				{
					Canvas.SetClipReplace(PaintClipRect);  //This will improve pdf export a lot, since we can now join all the BT/ET tags inside one, and then we can also join fonts inside the ET/BT tags.
				}
				else
				{
					if (Alpha == 0 || Alpha == 90)
					{
						Canvas.SetClipReplace(FinalRect);
					}
					else
					{
						Canvas.SetClipReplace(RectangleF.FromLTRB(PaintClipRect.Left, FinalRect.Top, PaintClipRect.Right, FinalRect.Bottom)); //rotated text can move to other used cells horizontally.
					}
				}

				real AcumDy = 0;
				for (int i = 0; i < TextLinesCount; i++)
				{
					TXRichString TextLine = TextLines[i];
					if ((Alpha == 0) &&
						(HJustify == THFlxAlignment.justify && TextLine.Split)
						|| (HJustify == THFlxAlignment.distributed)
						)
					{
						Canvas.SetClipReplace(FinalRect);
						WriteJustText(Workbook, Canvas, FontCache, Zoom100, ReverseRightToLeftStrings, 
                            AFont, DrawFontColor, Y[i] + AcumDy, SubData.Offset(Canvas, AFont), CellRect, Clp, TextLine.s,
                            TextLine.XExtent, MaxDescent[i], HJustify == THFlxAlignment.distributed, TextLine.AdaptFormat);
					}
					else
					{
						WriteText(Workbook, Canvas, FontCache, Zoom100, AFont, DrawFontColor, X[i],
							Y[i] + AcumDy, SubData.Offset(Canvas, AFont), GetVisualString(TextLine.s, ReverseRightToLeftStrings), 
                            Alpha, MaxDescent[i], TextLine.AdaptFormat);
					}
					AcumDy += dy;
				}
			}
            
		}
		#endregion

	}
	#endregion
}
