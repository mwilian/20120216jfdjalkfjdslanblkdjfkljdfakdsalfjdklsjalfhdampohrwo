using System;
using System.Text;

using FlexCel.Core;

#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
using Font = MonoTouch.UIKit.UIFont;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	using SizeF = System.Windows.Size;
	using real = System.Double;
	using System.Windows.Media;
	#else
	using real = System.Single;
	using System.Drawing;
	using System.Drawing.Imaging;
	#endif
#endif

namespace FlexCel.Render
{
	/// <summary>
	/// Static class to draw WordArt objects.
	/// </summary>
	internal sealed class DrawWordArt
	{
		private DrawWordArt(){}

        #region Utilities
        internal static Brush GetBrush(RectangleF Coords, TShapeProperties ShProp, ExcelFile Workbook, TShadowInfo ShadowInfo, float Zoom100)
        {
            return DrawShape.GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100);
        }

        internal static Pen GetPen(TShapeProperties ShProp, ExcelFile Workbook, TShadowInfo ShadowInfo)
        {
            return DrawShape.GetPen(ShProp, Workbook, ShadowInfo);
        }

        internal static string GetGeoText(TShapeProperties ShProp)
        {
            if (ShProp.ShapeOptions == null) return string.Empty;

            string UnicodeText = ShProp.ShapeOptions.AsUnicodeString(TShapeOption.gtextUNICODE, null);
            if (UnicodeText != null)
            {
                return UnicodeText;
            }

            string RTFText = ShProp.ShapeOptions.AsUnicodeString(TShapeOption.gtextRTF, null);
            if (RTFText != null)
            {
                return RTFText;
            }

            return String.Empty;            
        }

        internal static Font GetGeoFont(TShapeProperties ShProp)
        {
            if (ShProp.ShapeOptions == null) return new Font("Arial", 10);
            real Size =  ShProp.ShapeOptions.As1616(TShapeOption.gtextSize, 36);
            string FontName = ShProp.ShapeOptions.AsUnicodeString(TShapeOption.gtextFont, "Arial");

            FontStyle Style = FontStyle.Regular;

			if (ShProp.ShapeOptions.AsBool(TShapeOption.gtextFStrikethrough, false, 5)) Style |= FontStyle.Bold;
			if (ShProp.ShapeOptions.AsBool(TShapeOption.gtextFStrikethrough, false, 4)) Style |= FontStyle.Italic;
			
			/*  UnderLine and StrikeOut are not supported by Excel <= 2003. Even if the file has those bits set (Excel 2007 saves them)
			 * 
			if (ShProp.ShapeOptions.AsBool(TShapeOption.gtextFStrikethrough, false, 3)) Style |= FontStyle.Underline;			
			if (ShProp.ShapeOptions.AsBool(TShapeOption.gtextFStrikethrough, false, 0)) Style |= FontStyle.Strikeout;
			*/
			
            return ExcelFont.CreateFont(FontName, Size, Style);
            
        }
        #endregion

        internal static void DrawPlainText(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, float Zoom100)
        {
            string Text = GetGeoText(ShProp);
            if (Text == null) return;
			Text = Text.Replace("\n",String.Empty);
            string[] Lines = Text.Split('\r');
            if (Lines == null || Lines.Length <= 0) return;
			int LinesLength = Lines[Lines.Length - 1].Length == 0? Lines.Length - 1: Lines.Length;   //Last line is an empty enter.
			if (LinesLength <= 0) return;

            using (Font TextFont = GetGeoFont(ShProp))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.SaveTransform();
                    try
                    {
                        float LineGap = Canvas.FontLinespacing(TextFont);
                        SizeF[] Sizes = new SizeF[LinesLength];
                        Sizes[0] = Canvas.MeasureStringEmptyHasHeight(Lines[0], TextFont);
                        Sizes[0].Height -= LineGap; //Linespacing is not included here.
                            
                        SizeF sz = Sizes[0];
                        for (int i = 1; i < LinesLength; i++)
                        {
                            Sizes[i] = Canvas.MeasureStringEmptyHasHeight(Lines[i], TextFont);
                            if (Sizes[i].Width > sz.Width) sz.Width = Sizes[i].Width;
                            sz.Height += Sizes[i].Height;
                        }

                        if (sz.Width <=0 || sz.Height <= 0 || Coords.Width <=0 || Coords.Height <= 0) return;
                        float rx = Coords.Width / sz.Width;
                        float ry = Coords.Height / sz.Height;
                        Canvas.Scale(rx, ry);

                        using (Brush br = GetBrush(new RectangleF(Coords.Left/rx, Coords.Top/ry, sz.Width, sz.Height), ShProp, Workbook, ShadowInfo, Zoom100)) //Mast be selected AFTER scaling, so gradients work.
                        {
                            float y = LineGap;
                            for (int i = 0; i < LinesLength; i++)
                            {
                                y+=Sizes[i].Height;
                                float x = (sz.Width - Sizes[i].Width)/2f;
                                Canvas.DrawString(Lines[i], TextFont, pe, br,  Coords.Left / rx + x, Coords.Top / ry + y);                           
                            }
                        }
                    }
                    finally
                    {
                        Canvas.ResetTransform();
                    }

                }
                           
            }
        }

	}
}
